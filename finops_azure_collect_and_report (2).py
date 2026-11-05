#!/usr/bin/env python3
"""
Azure FinOps collector + analytics report
=========================================
Pulls cost data for EVERY subscription you can see (one Cost Management query
per subscription, so it works for EA, Dev/Test and Pay-As-You-Go subscriptions
alike), pulls Advisor cost recommendations, and builds an Excel report.

Needs only read access you already have (Reader/Contributor on the subscriptions).
It never changes anything in Azure.

Setup (once):
    pip install requests pandas openpyxl
    az login            # sign in with the account that has the 73 subscriptions

Run:
    python finops_azure_collect_and_report.py                  # 13 months, ActualCost
    python finops_azure_collect_and_report.py --cost-type AmortizedCost
    python finops_azure_collect_and_report.py --max-subs 3     # quick test first
    python finops_azure_collect_and_report.py --report-only    # rebuild Excel from cache

Output folder (default ./finops_output):
    FinOps_Report.xlsx        <- the analytics report
    cache/                    <- raw results, so re-runs do not hit Azure again
    errors.csv                <- subscriptions/queries that failed, with the reason
"""
import argparse
import calendar
import json
import os
import shutil
import subprocess
import sys
import time
from datetime import date, datetime, timedelta

import pandas as pd
import requests

ARM = "https://management.azure.com"
COST_API = "2023-11-01"
SUBS_API = "2022-12-01"
ARG_API = "2021-03-01"

# Rough mapping from Advisor resource type to Cost Management ServiceName,
# used only to compare Advisor savings with spend (best effort).
TYPE_TO_SERVICE = {
    "microsoft.compute/virtualmachines": "Virtual Machines",
    "microsoft.compute/virtualmachinescalesets": "Virtual Machines",
    "microsoft.compute/disks": "Storage",
    "microsoft.storage/storageaccounts": "Storage",
    "microsoft.sql/servers": "SQL Database",
    "microsoft.sql/servers/databases": "SQL Database",
    "microsoft.sql/managedinstances": "SQL Managed Instance",
    "microsoft.web/serverfarms": "Azure App Service",
    "microsoft.web/sites": "Azure App Service",
    "microsoft.cache/redis": "Redis Cache",
    "microsoft.documentdb/databaseaccounts": "Azure Cosmos DB",
    "microsoft.containerservice/managedclusters": "Azure Kubernetes Service",
    "microsoft.machinelearningservices/workspaces": "Azure Machine Learning",
    "microsoft.fabric/capacities": "Microsoft Fabric",
    "microsoft.network/publicipaddresses": "Virtual Network",
}


# --------------------------------------------------------------------------- #
# HTTP / auth
# --------------------------------------------------------------------------- #
class Client:
    """Thin ARM client: az-cli token, retry on 429/5xx, honours Retry-After."""

    def __init__(self):
        self.s = requests.Session()
        self._token = None
        self._token_time = 0

    def _get_token(self):
        if self._token and time.time() - self._token_time < 1500:
            return self._token
        az = shutil.which("az")
        if not az:
            sys.exit("Azure CLI ('az') not found. Install it and run 'az login'.")
        cmd = [az, "account", "get-access-token", "--resource", ARM,
               "--query", "accessToken", "-o", "tsv"]
        res = subprocess.run(cmd, capture_output=True, text=True, shell=(os.name == "nt"))
        if res.returncode != 0 or not res.stdout.strip():
            sys.exit("Could not get a token. Run 'az login' first.\n" + res.stderr)
        self._token = res.stdout.strip()
        self._token_time = time.time()
        return self._token

    def request(self, method, url, body=None, max_tries=8):
        for attempt in range(1, max_tries + 1):
            headers = {"Authorization": "Bearer " + self._get_token(),
                       "Content-Type": "application/json"}
            try:
                r = self.s.request(method, url, headers=headers,
                                   data=json.dumps(body) if body is not None else None,
                                   timeout=120)
            except requests.RequestException as e:
                wait = min(60, 2 ** attempt)
                print(f"    network error ({e}); retry in {wait}s")
                time.sleep(wait)
                continue
            if r.status_code == 401:
                self._token = None            # refresh token and retry
                continue
            if r.status_code in (429, 500, 502, 503, 504):
                wait = None
                for h in ("x-ms-ratelimit-microsoft.costmanagement-entity-retry-after",
                          "x-ms-ratelimit-microsoft.costmanagement-clienttype-retry-after",
                          "x-ms-ratelimit-microsoft.costmanagement-tenant-retry-after",
                          "Retry-After"):
                    if h in r.headers:
                        try:
                            wait = int(float(r.headers[h])) + 1
                            break
                        except ValueError:
                            pass
                wait = wait or min(90, 5 * attempt)
                print(f"    HTTP {r.status_code}; waiting {wait}s (try {attempt}/{max_tries})")
                time.sleep(wait)
                continue
            return r
        raise RuntimeError(f"Gave up after {max_tries} tries: {method} {url}")


# --------------------------------------------------------------------------- #
# Date helpers
# --------------------------------------------------------------------------- #
def first_of_month(d):
    return d.replace(day=1)


def add_months(d, n):
    y, m = divmod(d.year * 12 + (d.month - 1) + n, 12)
    return date(y, m + 1, 1)


def month_windows(months, today, chunk=6):
    """Split [start of (months-1) ago .. today] into windows of <= chunk months."""
    start = add_months(first_of_month(today), -(months - 1))
    wins, cur = [], start
    while cur <= today:
        nxt = add_months(cur, chunk)
        end = min(today, nxt - timedelta(days=1))
        wins.append((cur, end))
        cur = nxt
    return wins


# --------------------------------------------------------------------------- #
# Collection
# --------------------------------------------------------------------------- #
def list_subscriptions(c):
    url = f"{ARM}/subscriptions?api-version={SUBS_API}"
    out = []
    while url:
        r = c.request("GET", url)
        r.raise_for_status()
        j = r.json()
        for s in j.get("value", []):
            out.append({
                "SubscriptionId": s["subscriptionId"],
                "SubscriptionName": s.get("displayName", ""),
                "State": s.get("state", ""),
                "Offer": (s.get("subscriptionPolicies") or {}).get("quotaId", ""),
            })
        url = j.get("nextLink")
    return pd.DataFrame(out)


def cost_query(c, sub_id, d_from, d_to, cost_type, granularity, grouping):
    """Run one Cost Management query (follows nextLink). Returns DataFrame."""
    body = {
        "type": cost_type,
        "timeframe": "Custom",
        "timePeriod": {"from": f"{d_from.isoformat()}T00:00:00Z",
                       "to": f"{d_to.isoformat()}T23:59:59Z"},
        "dataset": {
            "granularity": granularity,
            "aggregation": {"totalCost": {"name": "Cost", "function": "Sum"}},
            "grouping": [{"type": "Dimension", "name": g} for g in grouping],
        },
    }
    url = (f"{ARM}/subscriptions/{sub_id}/providers/Microsoft.CostManagement/query"
           f"?api-version={COST_API}")
    frames = []
    while url:
        r = c.request("POST", url, body)
        if r.status_code != 200:
            raise RuntimeError(f"HTTP {r.status_code}: {r.text[:400]}")
        p = r.json().get("properties", {})
        cols = [x["name"] for x in p.get("columns", [])]
        rows = p.get("rows", [])
        if rows:
            frames.append(pd.DataFrame(rows, columns=cols))
        url = p.get("nextLink")
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def normalize_month(df):
    """Monthly queries return UsageDate (yyyymmdd int) or BillingMonth; make Month."""
    for col in ("UsageDate", "BillingMonth"):
        if col in df.columns:
            s = df[col].astype(str)
            if col == "UsageDate":
                m = pd.to_datetime(s, format="%Y%m%d", errors="coerce")
            else:
                m = pd.to_datetime(s, errors="coerce")
            df["Month"] = m.dt.to_period("M").dt.to_timestamp()
            return df.drop(columns=[col])
    return df


QUERIES = {
    # name: (granularity, grouping, how many months from the end to cover)
    "svc_rg": ("Monthly", ["ServiceName", "ResourceGroupName"], None),
    "pricing": ("Monthly", ["PricingModel", "ServiceName"], None),
    "resources": ("None", ["ResourceId", "ServiceName"], 1),   # last full month only
}


def collect_cost(c, subs, out_dir, months, cost_type, today, pause):
    cache = os.path.join(out_dir, "cache")
    os.makedirs(cache, exist_ok=True)
    errors = []
    wins = month_windows(months, today)
    last_full_start = add_months(first_of_month(today), -1)
    last_full_end = first_of_month(today) - timedelta(days=1)

    for i, row in enumerate(subs.itertuples(index=False), 1):
        sid = row.SubscriptionId
        print(f"[{i}/{len(subs)}] {row.SubscriptionName} ({row.Offer})")
        for qname, (gran, grouping, last_n) in QUERIES.items():
            f = os.path.join(cache, f"{qname}__{cost_type}__{sid}.csv")
            if os.path.exists(f):
                continue
            try:
                if last_n:
                    frames = [cost_query(c, sid, last_full_start, last_full_end,
                                         cost_type, gran, grouping)]
                else:
                    frames = []
                    for (a, b) in wins:
                        try:
                            frames.append(cost_query(c, sid, a, b, cost_type, gran, grouping))
                        except RuntimeError as e:
                            # oldest window can fall outside the 13-month retention
                            if a == wins[0][0] and "400" in str(e):
                                a2 = add_months(a, 1)
                                print(f"    oldest window rejected, retrying from {a2}")
                                frames.append(cost_query(c, sid, a2, b, cost_type, gran, grouping))
                            else:
                                raise
                df = pd.concat([x for x in frames if len(x)], ignore_index=True) if any(len(x) for x in frames) else pd.DataFrame()
                if len(df) and gran == "Monthly":
                    df = normalize_month(df)
                df.to_csv(f, index=False)
            except Exception as e:      # keep going; record the reason
                msg = str(e).replace("\n", " ")[:500]
                print(f"    {qname} failed: {msg}")
                errors.append({"SubscriptionId": sid, "SubscriptionName": row.SubscriptionName,
                               "Offer": row.Offer, "Query": qname, "Error": msg})
            time.sleep(pause)
    return errors


def arg_query(c, sub_ids, query):
    """Run an Azure Resource Graph query across subscriptions (handles paging)."""
    rows, token = [], None
    url = f"{ARM}/providers/Microsoft.ResourceGraph/resources?api-version={ARG_API}"
    while True:
        opts = {"$top": 1000, "resultFormat": "objectArray"}
        if token:
            opts["$skipToken"] = token
        r = c.request("POST", url, {"subscriptions": sub_ids, "query": query, "options": opts})
        if r.status_code != 200:
            raise RuntimeError(f"ARG HTTP {r.status_code}: {r.text[:300]}")
        j = r.json()
        rows += j.get("data", [])
        token = j.get("$skipToken")
        if not token:
            break
    return rows


def collect_advisor(c, sub_ids, out_dir):
    f = os.path.join(out_dir, "cache", "advisor.csv")
    if os.path.exists(f):
        return
    q = """advisorresources
| where type == 'microsoft.advisor/recommendations'
| where properties.category == 'Cost'
| extend _rm = tostring(properties.resourceMetadata.resourceId),
         _ix = indexof(tolower(id), '/providers/microsoft.advisor/recommendations')
| extend resourceId = iff(isnotempty(_rm), _rm,
                          iff(_ix > 0 and tolower(substring(id, 0, _ix)) contains '/resourcegroups/', substring(id, 0, _ix), ''))
| extend impacted = tolower(tostring(properties.impactedField)),
         problem = tostring(properties.shortDescription.problem),
         solution = tostring(properties.shortDescription.solution),
         impact = tostring(properties.impact),
         annualSavings = todouble(properties.extendedProperties.annualSavingsAmount),
         monthlySavings = todouble(properties.extendedProperties.savingsAmount),
         currency = tostring(properties.extendedProperties.savingsCurrency),
         sku = coalesce(tostring(properties.extendedProperties.sku), tostring(properties.extendedProperties.vmSize)),
         region = tostring(properties.extendedProperties.region),
         term = tostring(properties.extendedProperties.term),
         qty = tostring(properties.extendedProperties.qty),
         lookBack = tostring(properties.extendedProperties.lookBackPeriod),
         impactedValue = tostring(properties.impactedValue),
         lastUpdated = tostring(properties.lastUpdated)
| project subscriptionId, resourceGroup, impacted, resourceId, problem, solution, impact,
          annualSavings, monthlySavings, currency, sku, region, term, qty, lookBack,
          impactedValue, lastUpdated"""
    try:
        rows = arg_query(c, sub_ids, q)
    except RuntimeError as e:
        print("Advisor query failed:", e)
        pd.DataFrame().to_csv(f, index=False)
        return
    pd.DataFrame(rows).to_csv(f, index=False)
    print(f"Advisor: {len(rows)} cost recommendations")


def collect_inventory(c, sub_ids, out_dir):
    """Compute/storage/database inventory used to map commitments to real resources."""
    f = os.path.join(out_dir, "cache", "inventory.csv")
    if os.path.exists(f):
        return
    q = """resources
| where type in~ ('microsoft.compute/virtualmachines','microsoft.compute/virtualmachinescalesets',
                  'microsoft.web/serverfarms','microsoft.storage/storageaccounts',
                  'microsoft.sql/servers/databases','microsoft.sql/managedinstances',
                  'microsoft.dbforpostgresql/flexibleservers','microsoft.dbformysql/flexibleservers')
| project id, name, type = tolower(type), subscriptionId, resourceGroup, location,
          skuName = iff(type =~ 'microsoft.compute/virtualmachines', tostring(properties.hardwareProfile.vmSize), tostring(sku.name)),
          skuTier = tostring(sku.tier), skuCapacity = tostring(sku.capacity),
          accessTier = tostring(properties.accessTier),
          power = tostring(properties.extended.instanceView.powerState.displayStatus),
          owner = tostring(tags['Owner']), environment = tostring(tags['Environment'])"""
    try:
        rows = arg_query(c, sub_ids, q)
    except RuntimeError as e:
        print("Inventory query failed:", e)
        pd.DataFrame().to_csv(f, index=False)
        return
    pd.DataFrame(rows).to_csv(f, index=False)
    print(f"Inventory: {len(rows)} resources")


# --------------------------------------------------------------------------- #
# Loading cached data
# --------------------------------------------------------------------------- #
def load_cached(out_dir, cost_type, subs):
    cache = os.path.join(out_dir, "cache")
    meta = subs.set_index("SubscriptionId")[["SubscriptionName", "Offer", "State"]]
    data = {}
    for qname in QUERIES:
        frames = []
        for sid in subs["SubscriptionId"]:
            f = os.path.join(cache, f"{qname}__{cost_type}__{sid}.csv")
            if os.path.exists(f) and os.path.getsize(f) > 2:
                d = pd.read_csv(f)
                if len(d):
                    d["SubscriptionId"] = sid
                    frames.append(d)
        df = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
        if len(df):
            if "Month" in df.columns:
                df["Month"] = pd.to_datetime(df["Month"])
            df = df.join(meta, on="SubscriptionId")
            if "Cost" not in df.columns and "PreTaxCost" in df.columns:
                df = df.rename(columns={"PreTaxCost": "Cost"})
            df["Environment"] = df["Offer"].fillna("").str.startswith("MSDNDevTest").map(
                {True: "Dev/Test offer", False: "Other offers"})
        data[qname] = df
    inv_f = os.path.join(cache, "inventory.csv")
    inv = pd.read_csv(inv_f) if os.path.exists(inv_f) and os.path.getsize(inv_f) > 2 else pd.DataFrame()
    data["inventory"] = inv
    adv_f = os.path.join(cache, "advisor.csv")
    adv = pd.read_csv(adv_f) if os.path.exists(adv_f) and os.path.getsize(adv_f) > 2 else pd.DataFrame()
    if len(adv):
        adv = adv.join(meta, on="subscriptionId")
        adv["annualSavings"] = adv["annualSavings"].fillna(adv["monthlySavings"] * 12)
        adv["Service"] = adv["impacted"].str.lower().map(TYPE_TO_SERVICE).fillna(adv["impacted"])
    return data, adv



# --------------------------------------------------------------------------- #
# Commitment (Reservation / Savings Plan) -> actual resources
# --------------------------------------------------------------------------- #
VM = "microsoft.compute/virtualmachines"
VMSS = "microsoft.compute/virtualmachinescalesets"
ASP = "microsoft.web/serverfarms"
SA = "microsoft.storage/storageaccounts"
SQLDB = "microsoft.sql/servers/databases"
SQLMI = "microsoft.sql/managedinstances"
PG = "microsoft.dbforpostgresql/flexibleservers"
MY = "microsoft.dbformysql/flexibleservers"
PREMIUM_ASP = ("premiumv3", "premiumv2", "isolatedv2", "premium0v3")


def solution_types(solution):
    s = (solution or "").lower()
    if "savings plan" in s:
        return [VM, VMSS, ASP], "Savings plan: eligible compute in subscription"
    if "blob storage reserved" in s:
        return [SA], "Blob reserved capacity: storage accounts"
    if "virtual machine reserved" in s or "vm reserved" in s:
        return [VM], "VM reserved instance"
    if "sql database reserved" in s or "sql paas db reserved" in s:
        return [SQLDB], "SQL Database reserved capacity"
    if "managed instance" in s and "reserved" in s:
        return [SQLMI], "SQL Managed Instance reserved capacity"
    if "postgresql reserved" in s:
        return [PG], "PostgreSQL reserved capacity"
    if "mysql reserved" in s:
        return [MY], "MySQL reserved capacity"
    if "app service reserved" in s:
        return [ASP], "App Service reserved capacity"
    return None, ""


def norm_region(x):
    return str(x or "").replace(" ", "").lower()


def build_commitment_mapping(adv, inv, res_cost):
    """One row per (commitment recommendation x real resource), with an
    allocation of the recommendation's savings by each resource's share of cost.
    The allocation is an ESTIMATE: Azure does not attribute commitments to resources."""
    if adv.empty or inv.empty:
        return pd.DataFrame()
    a = adv.copy()
    for col in ("solution", "sku", "region", "term", "qty", "lastUpdated", "lookBack"):
        if col not in a.columns:
            a[col] = ""
    a["solution"] = a["solution"].fillna("")
    a = a[a["solution"].str.lower().str.contains("reserv|savings plan", regex=True)]
    a = a[a["resourceId"].fillna("") == ""]          # subscription-level commitments only
    if a.empty:
        return pd.DataFrame()
    a = a.sort_values("lastUpdated").drop_duplicates(
        ["subscriptionId", "solution", "sku", "region", "term"], keep="last")
    inv = inv.copy()
    inv["regionKey"] = inv["location"].map(norm_region)
    inv["idKey"] = inv["id"].str.lower()
    cost_map = res_cost if res_cost is not None else {}
    out = []
    for r in a.itertuples(index=False):
        types, label = solution_types(r.solution)
        annual = r.annualSavings if pd.notna(r.annualSavings) else (r.monthlySavings * 12 if pd.notna(r.monthlySavings) else 0.0)
        base = {"Subscription": getattr(r, "SubscriptionName", ""), "Offer": getattr(r, "Offer", ""),
                "Recommendation": r.solution, "Commitment type": label or "Not mapped",
                "Term": r.term, "Region": r.region, "SKU": r.sku, "Recommended qty": r.qty,
                "Look-back": r.lookBack, "Annual savings, whole recommendation": annual,
                "Currency": r.currency}
        if not types:
            out.append({**base, "Match quality": "Not mapped (commitment type not recognised)"})
            continue
        c = inv[(inv["subscriptionId"] == r.subscriptionId) & (inv["type"].isin(types))]
        quality = "Possible (same type in subscription)"
        if r.region:
            c2 = c[c["regionKey"] == norm_region(r.region)]
            if len(c2):
                c, quality = c2, "Possible (same type and region)"
        if r.sku and VM in types and len(types) == 1:
            c2 = c[c["skuName"].str.lower() == str(r.sku).lower()]
            if len(c2):
                c, quality = c2, "Exact (same VM size and region)"
        if ASP in types:                         # only premium/isolated plans are covered
            c = c[(c["type"] != ASP) | (c["skuTier"].str.lower().str.replace(" ", "").isin(PREMIUM_ASP))]
        if "savings plan" in label.lower():
            c = c[(c["type"] != VM) | (c["power"].fillna("").str.lower().str.contains("running") | c["power"].isna() | (c["power"] == ""))]
        if c.empty:
            out.append({**base, "Match quality": "No matching resource found in inventory"})
            continue
        c = c.copy()
        c["cost"] = c["idKey"].map(cost_map).fillna(0.0)
        pw = c["power"].fillna("").str.lower()
        c["notrun"] = (c["type"] == VM) & (pw != "") & ~pw.str.contains("running")
        c["weight"] = c["cost"].where(~c["notrun"], 0.0)
        tot = c["weight"].sum()
        if tot > 0:
            c["share"] = c["weight"] / tot
        else:
            ok = ~c["notrun"]
            c["share"] = ok.astype(float) / ok.sum() if ok.sum() else 0.0
        for x in c.sort_values("cost", ascending=False).head(500).itertuples(index=False):
            out.append({**base, "Match quality": quality, "Resource group": x.resourceGroup,
                        "Resource name": x.name, "Resource type": x.type.split("/")[-1],
                        "Resource SKU": x.skuName, "Power state": x.power, "Owner": x.owner,
                        "Environment": x.environment, "Last-month cost": round(x.cost, 2),
                        "Cost share %": x.share,
                        "Note": "Not running: excluded from allocation" if x.notrun else "",
                        "Allocated annual savings (estimate)": round(annual * x.share, 2),
                        "Resource ID": x.id})
    df = pd.DataFrame(out)
    if len(df):
        df = df.sort_values(["Subscription", "Recommendation", "Term", "Last-month cost"],
                            ascending=[True, True, True, False])
    return df

# --------------------------------------------------------------------------- #
# Analytics
# --------------------------------------------------------------------------- #
def pivot_months(df, index, value="Cost"):
    p = df.groupby([index, "Month"])[value].sum().unstack("Month", fill_value=0)
    p = p.reindex(sorted(p.columns), axis=1)
    p["Total"] = p.sum(axis=1)
    return p.sort_values("Total", ascending=False)


def add_trend_cols(p, last_full):
    """Add share, last-3-month average, MoM % to a month pivot."""
    months = [c for c in p.columns if isinstance(c, pd.Timestamp)]
    full = [m for m in months if m <= last_full]
    out = p.copy()
    out["Share %"] = out["Total"] / out["Total"].sum() if out["Total"].sum() else 0
    out["Cumulative %"] = out["Share %"].cumsum()
    if len(full) >= 2:
        last, prev = full[-1], full[-2]
        out["MoM % (last full month)"] = (out[last] - out[prev]) / out[prev].replace(0, pd.NA)
    if len(full) >= 3:
        out["Avg last 3 full months"] = out[full[-3:]].mean(axis=1)
    out.columns = [c.strftime("%Y-%m") if isinstance(c, pd.Timestamp) else c for c in out.columns]
    return out


def build_report(out_dir, cost_type, subs, errors_df, today):
    data, adv = load_cached(out_dir, cost_type, subs)
    svc = data["svc_rg"]
    if svc.empty:
        sys.exit("No cost data was collected. Check errors.csv (usually a permission or billing-visibility issue).")
    for col in ("ServiceName", "ResourceGroupName"):
        svc[col] = svc[col].fillna("(none)")
    currency = ", ".join(sorted(svc["Currency"].dropna().unique())) if "Currency" in svc else ""

    months_all = sorted(svc["Month"].unique())
    cur_month = pd.Timestamp(first_of_month(today))
    full_months = [m for m in months_all if m < cur_month]
    last_full = full_months[-1] if full_months else months_all[-1]
    mtd = svc.loc[svc["Month"] == cur_month, "Cost"].sum()

    sheets = {}
    by_service = add_trend_cols(pivot_months(svc, "ServiceName"), last_full)
    by_sub_raw = svc.assign(Sub=svc["SubscriptionName"] + " [" + svc["Offer"].fillna("") + "]")
    by_sub = add_trend_cols(pivot_months(by_sub_raw, "Sub"), last_full)
    by_offer = add_trend_cols(pivot_months(svc.assign(Offer2=svc["Offer"].fillna("(unknown)")), "Offer2"), last_full)
    by_env = add_trend_cols(pivot_months(svc, "Environment"), last_full)
    by_rg = add_trend_cols(pivot_months(svc.assign(RG=svc["SubscriptionName"] + " / " + svc["ResourceGroupName"]), "RG"), last_full).head(200)

    # Service x subscription for last full month (where is each service's money)
    lf = svc[svc["Month"] == last_full]
    svc_sub_last = lf.pivot_table(index="ServiceName", columns="SubscriptionName", values="Cost",
                                  aggfunc="sum", fill_value=0)
    svc_sub_last["Total"] = svc_sub_last.sum(axis=1)
    svc_sub_last = svc_sub_last.sort_values("Total", ascending=False).head(25)

    # Anomalies: service x subscription, last full month vs avg of 3 prior months
    anomalies = pd.DataFrame()
    if len(full_months) >= 4:
        prior = full_months[-4:-1]
        g = svc.groupby(["SubscriptionName", "ServiceName", "Month"])["Cost"].sum().unstack("Month", fill_value=0)
        a = pd.DataFrame({"Last full month": g.get(last_full, 0),
                          "Avg prior 3 months": g[[m for m in prior if m in g.columns]].mean(axis=1)})
        a["Change $"] = a["Last full month"] - a["Avg prior 3 months"]
        a["Change %"] = a["Change $"] / a["Avg prior 3 months"].replace(0, pd.NA)
        anomalies = a[(a["Change $"] > 500) & ((a["Change %"] > 0.30) | a["Change %"].isna())] \
            .sort_values("Change $", ascending=False).head(50).reset_index()

    # Pricing model / commitment coverage
    pricing = data["pricing"]
    cov = pd.DataFrame()
    if not pricing.empty and "PricingModel" in pricing.columns:
        pricing["PricingModel"] = pricing["PricingModel"].fillna("Unknown")
        cov = pricing.groupby(["Month", "PricingModel"])["Cost"].sum().unstack("PricingModel", fill_value=0)
        cov["Total"] = cov.sum(axis=1)
        committed = [c for c in cov.columns if c in ("Reservation", "SavingsPlan")]
        cov["Commitment-covered %"] = cov[committed].sum(axis=1) / cov["Total"].replace(0, pd.NA) if committed else 0
        cov.index = cov.index.strftime("%Y-%m")
        cov = cov.reset_index()

    # Top resources (last full month)
    res = data["resources"]
    top_res = pd.DataFrame()
    if not res.empty:
        top_res = res.sort_values("Cost", ascending=False).head(300)[
            ["SubscriptionName", "Offer", "ServiceName", "ResourceId", "Cost", "Currency"]]

    # Advisor
    adv_svc = adv_sub = adv_top = savings_vs_spend = pd.DataFrame()
    adv_total = 0.0
    if not adv.empty:
        adv_total = adv["annualSavings"].sum()
        adv_svc = adv.groupby("Service")["annualSavings"].agg(["sum", "count"]).sort_values("sum", ascending=False)
        adv_svc.columns = ["Annual savings (potential)", "Recommendations"]
        adv_svc["Share %"] = adv_svc.iloc[:, 0] / adv_svc.iloc[:, 0].sum()
        adv_sub = adv.groupby(["SubscriptionName", "Offer"])["annualSavings"].agg(["sum", "count"]).sort_values("sum", ascending=False)
        adv_sub.columns = ["Annual savings (potential)", "Recommendations"]
        adv_sub = adv_sub.reset_index()
        adv_top = adv.sort_values("annualSavings", ascending=False).head(100)[
            ["SubscriptionName", "resourceGroup", "Service", "problem", "solution", "impact",
             "annualSavings", "currency", "resourceId"]]
        # Savings vs annualised spend per service (best-effort name match)
        n_full = min(12, len(full_months))
        spend12 = svc[svc["Month"].isin(full_months[-n_full:])].groupby("ServiceName")["Cost"].sum()
        sv = adv_svc.iloc[:, [0]].join(spend12.rename(f"Spend last {n_full} full months"), how="left")
        sv["Savings as % of spend"] = sv.iloc[:, 0] / sv.iloc[:, 1]
        savings_vs_spend = sv.reset_index()

    # Commitment recommendations mapped to real resources (estimate)
    commit_map = pd.DataFrame()
    inv = data.get("inventory", pd.DataFrame())
    if not adv.empty and not inv.empty:
        adv_m = adv.join(subs.set_index("SubscriptionId")[["SubscriptionName", "Offer"]], on="subscriptionId", rsuffix="_s") \
            if "SubscriptionName" not in adv.columns else adv
        res_cost = None
        if not top_res.empty or not data["resources"].empty:
            rc = data["resources"]
            res_cost = rc.assign(k=rc["ResourceId"].str.lower()).groupby("k")["Cost"].sum().to_dict()
        commit_map = build_commitment_mapping(adv_m, inv, res_cost)

    # Summary KPIs
    n_full = min(12, len(full_months))
    last12 = svc[svc["Month"].isin(full_months[-n_full:])]["Cost"].sum() if full_months else 0
    prev = full_months[-2] if len(full_months) >= 2 else None
    lf_total = svc.loc[svc["Month"] == last_full, "Cost"].sum()
    prev_total = svc.loc[svc["Month"] == prev, "Cost"].sum() if prev is not None else None
    top5 = by_service.head(5)
    kpis = [
        ("Cost type", cost_type),
        ("Note", "'Total' and 'Share %' columns include the current partial month"),
        ("Currency", currency),
        ("Report date", today.isoformat()),
        ("Subscriptions listed / with cost data", f"{len(subs)} / {svc['SubscriptionId'].nunique()}"),
        ("Subscriptions with failed queries", errors_df["SubscriptionId"].nunique() if len(errors_df) else 0),
        ("Months of data", f"{months_all[0]:%Y-%m} to {months_all[-1]:%Y-%m}"),
        (f"Spend, last {n_full} full months", round(last12, 2)),
        (f"Last full month ({last_full:%Y-%m})", round(lf_total, 2)),
        ("Month-over-month change", (lf_total - prev_total) / prev_total if prev_total else "n/a"),
        (f"Current month to date ({cur_month:%Y-%m})", round(mtd, 2)),
        ("Top 5 services share of spend", float(top5["Share %"].sum())),
        ("Advisor potential annual savings (cost recs)", round(adv_total, 2)),
        ("Advisor savings as % of last-12-month spend",
         (adv_total / last12) if last12 and adv_total and n_full == 12 else "n/a (needs 12 full months)"),
    ]

    sheets["Monthly_by_Service"] = by_service.reset_index()
    sheets["Monthly_by_Subscription"] = by_sub.reset_index()
    sheets["By_Offer_Type"] = by_offer.reset_index()
    sheets["Dev-Test_vs_Other"] = by_env.reset_index()
    sheets["Top_ResourceGroups"] = by_rg.reset_index()
    sheets["Service_x_Subscription"] = svc_sub_last.reset_index()
    if len(anomalies):
        sheets["Anomalies"] = anomalies
    if len(cov):
        sheets["Commitment_Coverage"] = cov
    if len(top_res):
        sheets["Top_Resources_LastMonth"] = top_res
    if len(adv_svc):
        sheets["Advisor_by_Service"] = adv_svc.reset_index()
        sheets["Advisor_by_Subscription"] = adv_sub
        sheets["Advisor_Top_Recs"] = adv_top
        sheets["Savings_vs_Spend"] = savings_vs_spend
    if len(commit_map):
        sheets["Commitment_Resource_Mapping"] = commit_map
    subs_out = subs.copy()
    sheets["Subscriptions"] = subs_out
    if len(errors_df):
        sheets["Data_Quality_Errors"] = errors_df

    write_excel(os.path.join(out_dir, "FinOps_Report.xlsx"), kpis, sheets, by_service)
    print("\nReport written:", os.path.join(out_dir, "FinOps_Report.xlsx"))


# --------------------------------------------------------------------------- #
# Excel output
# --------------------------------------------------------------------------- #
def write_excel(path, kpis, sheets, by_service):
    from openpyxl.chart import BarChart, Reference
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        pd.DataFrame(kpis, columns=["KPI", "Value"]).to_excel(xw, sheet_name="Summary", index=False)
        for name, df in sheets.items():
            df.to_excel(xw, sheet_name=name[:31], index=False)
        wb = xw.book

        hdr_fill = PatternFill("solid", fgColor="0B1F3A")
        for ws in wb.worksheets:
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = hdr_fill
                cell.alignment = Alignment(wrap_text=True, vertical="center")
            ws.freeze_panes = "B2"
            for ci, col in enumerate(ws.columns, 1):
                header = str(col[0].value or "")
                width = min(60, max(12, max((len(str(c.value)) for c in col[:60] if c.value is not None), default=10) + 2))
                ws.column_dimensions[get_column_letter(ci)].width = width
                pct = "%" in header
                for c in col[1:]:
                    if isinstance(c.value, (int, float)):
                        c.number_format = "0.0%" if pct else "#,##0"
        ws = wb["Summary"]
        ws.column_dimensions["A"].width = 52
        ws.column_dimensions["B"].width = 32
        for row in ws.iter_rows(min_row=2):
            v = row[1].value
            if isinstance(v, float) and abs(v) < 5 and ("change" in str(row[0].value).lower() or "share" in str(row[0].value).lower() or "%" in str(row[0].value)):
                row[1].number_format = "0.0%"
            elif isinstance(v, (int, float)):
                row[1].number_format = "#,##0"
            row[1].alignment = Alignment(horizontal="left")

        # Chart 1: stacked monthly spend for top 8 services
        ws = wb["Monthly_by_Service"]
        month_cols = [i for i, c in enumerate(ws[1], 1)
                      if isinstance(c.value, str) and len(c.value) == 7 and c.value[4] == "-"]
        if month_cols:
            n = min(8, ws.max_row - 1)
            ch = BarChart()
            ch.type, ch.grouping, ch.overlap = "col", "stacked", 100
            ch.title = "Monthly spend, top 8 services"
            ch.height, ch.width = 11, 26
            data = Reference(ws, min_col=1, max_col=month_cols[-1], min_row=2, max_row=1 + n)
            # series from rows; first col = title, but month cols start at 2
            ch.add_data(data, from_rows=True, titles_from_data=True)
            ch.set_categories(Reference(ws, min_col=month_cols[0], max_col=month_cols[-1], min_row=1))
            ws.add_chart(ch, f"A{ws.max_row + 3}")

        # Chart 2: top 10 services by total
        ws = wb["Monthly_by_Service"]
        tot_col = [i for i, c in enumerate(ws[1], 1) if c.value == "Total"][0]
        ch2 = BarChart()
        ch2.type = "bar"
        ch2.title = "Top 10 services by total spend"
        ch2.height, ch2.width = 11, 18
        ch2.add_data(Reference(ws, min_col=tot_col, min_row=1, max_row=min(11, ws.max_row)), titles_from_data=True)
        ch2.set_categories(Reference(ws, min_col=1, min_row=2, max_row=min(11, ws.max_row)))
        ch2.legend = None
        ch2.x_axis.scaling.orientation = "maxMin"
        ws.add_chart(ch2, f"A{ws.max_row + 27}")


# --------------------------------------------------------------------------- #
def main():
    ap = argparse.ArgumentParser(description="Azure FinOps collector and report")
    ap.add_argument("--out", default="finops_output")
    ap.add_argument("--months", type=int, default=13, help="months incl. current (Azure keeps ~13)")
    ap.add_argument("--cost-type", default="ActualCost", choices=["ActualCost", "AmortizedCost"])
    ap.add_argument("--max-subs", type=int, default=0, help="limit subscriptions (for a quick test)")
    ap.add_argument("--pause", type=float, default=1.0, help="seconds between queries")
    ap.add_argument("--report-only", action="store_true", help="skip Azure calls; use cache")
    args = ap.parse_args()

    os.makedirs(os.path.join(args.out, "cache"), exist_ok=True)
    today = date.today()
    c = Client()

    subs_file = os.path.join(args.out, "cache", "subscriptions.csv")
    if args.report_only and os.path.exists(subs_file):
        subs = pd.read_csv(subs_file)
    else:
        subs = list_subscriptions(c)
        subs.to_csv(subs_file, index=False)
    print(f"{len(subs)} subscriptions found. Offers:\n{subs['Offer'].value_counts().to_string()}\n")
    work = subs.head(args.max_subs) if args.max_subs else subs

    errors = []
    if not args.report_only:
        errors = collect_cost(c, work, args.out, args.months, args.cost_type, today, args.pause)
        collect_advisor(c, list(work["SubscriptionId"]), args.out)
        collect_inventory(c, list(work["SubscriptionId"]), args.out)
        err_f = os.path.join(args.out, "errors.csv")
        old = pd.read_csv(err_f) if os.path.exists(err_f) else pd.DataFrame()
        errs = pd.concat([old, pd.DataFrame(errors)], ignore_index=True).drop_duplicates() if errors or len(old) else pd.DataFrame()
        errs.to_csv(err_f, index=False)
    err_f = os.path.join(args.out, "errors.csv")
    errors_df = pd.read_csv(err_f) if os.path.exists(err_f) and os.path.getsize(err_f) > 2 else pd.DataFrame()

    build_report(args.out, args.cost_type, work, errors_df, today)


if __name__ == "__main__":
    main()

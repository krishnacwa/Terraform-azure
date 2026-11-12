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
            "grouping": [({"type": "TagKey", "name": g[4:]} if g.startswith("tag:") else {"type": "Dimension", "name": g}) for g in grouping],
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
    "region": ("None", ["ResourceLocation"], 1),
}


def queries_for(tag_keys):
    q = dict(QUERIES)
    for k in tag_keys:
        q["tag_" + k] = ("None", ["tag:" + k], 1)
    return q


def collect_cost(c, subs, out_dir, months, cost_type, today, pause, tag_keys=()):
    cache = os.path.join(out_dir, "cache")
    os.makedirs(cache, exist_ok=True)
    errors = []
    wins = month_windows(months, today)
    last_full_start = add_months(first_of_month(today), -1)
    last_full_end = first_of_month(today) - timedelta(days=1)

    for i, row in enumerate(subs.itertuples(index=False), 1):
        sid = row.SubscriptionId
        print(f"[{i}/{len(subs)}] {row.SubscriptionName} ({row.Offer})")
        for qname, (gran, grouping, last_n) in queries_for(tag_keys).items():
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
def load_cached(out_dir, cost_type, subs, tag_keys=()):
    cache = os.path.join(out_dir, "cache")
    meta = subs.set_index("SubscriptionId")[["SubscriptionName", "Offer", "State"]]
    data = {}
    for qname in queries_for(tag_keys):
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


def build_report(out_dir, cost_type, subs, errors_df, today, tag_keys=(), growth=0.0, fy_history=None):
    data, adv = load_cached(out_dir, cost_type, subs, tag_keys)
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
    ctx = dict(today=today, cost_type=cost_type, currency=currency, svc=svc, by_service=by_service, by_sub=by_sub, by_offer=by_offer,
               by_env=by_env, by_rg=by_rg, full_months=full_months, cur_month=cur_month, last_full=last_full, mtd=mtd, n_full=n_full,
               last12=last12, lf_total=lf_total, prev_total=prev_total, anomalies=anomalies, pricing=data["pricing"], cov=cov,
               top_res=top_res, adv=adv, adv_svc=adv_svc, adv_top=adv_top, savings_vs_spend=savings_vs_spend, commit_map=commit_map,
               adv_total=adv_total, subs=subs, errors_df=errors_df, region=data.get("region", pd.DataFrame()),
               tags={k: data.get("tag_" + k, pd.DataFrame()) for k in tag_keys}, growth=growth, fy_history=fy_history or {})
    try:
        write_html(os.path.join(out_dir, "FinOps_Dashboard.html"), ctx)
        print("Dashboard written:", os.path.join(out_dir, "FinOps_Dashboard.html"))
    except Exception as e:                      # the Excel report is already saved
        print("HTML dashboard failed:", e)


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

# --------------------------------------------------------------------------- #
# HTML dashboard (single file, no libraries, opens in any browser)
# --------------------------------------------------------------------------- #
import html as _html


def _e(x):
    return _html.escape("" if x is None else str(x))


def _m(x, dec=0):
    try:
        if x is None or pd.isna(x):
            return "-"
        return f"{float(x):,.{dec}f}"
    except (TypeError, ValueError):
        return _e(x)


def _p(x, dec=1):
    try:
        if x is None or pd.isna(x):
            return "-"
        return f"{float(x):+.{dec}f}%"
    except (TypeError, ValueError):
        return "-"


def _kpi(label, value, sub=""):
    return (f'<div class="kpi"><div class="kl">{_e(label)}</div><div class="kv">{_e(value)}</div>'
            f'<div class="ks">{_e(sub)}</div></div>')


def _hbars(rows, total=None):
    rows = [(l, float(v)) for l, v in rows if pd.notna(v)]
    if not rows:
        return '<p class="muted">No data.</p>'
    mx = max(v for _, v in rows) or 1
    tot = total if total else sum(v for _, v in rows)
    out = ['<div class="bars">']
    for l, v in rows:
        w = max(0.5, 100 * v / mx) if v > 0 else 0
        share = f"{100 * v / tot:.1f}%" if tot else ""
        out.append(f'<div class="bar"><div class="lbl" title="{_e(l)}">{_e(l)}</div>'
                   f'<div class="track"><div class="fill" style="width:{w:.1f}%"></div></div>'
                   f'<div class="val">{_m(v)} <span class="sh">{share}</span></div></div>')
    out.append("</div>")
    return "".join(out)


def _columns(labels, values, notes=None):
    mx = max([v for v in values if pd.notna(v)] + [1])
    out = ['<div class="cols">']
    for i, (l, v) in enumerate(zip(labels, values)):
        h = 100 * v / mx if pd.notna(v) else 0
        note = f'<div class="cn">{_e(notes[i])}</div>' if notes else ""
        out.append(f'<div class="col"><div class="cv">{_m(v)}</div><div class="cb" style="height:{h * 0.7:.1f}%"></div>'
                   f'<div class="cl">{_e(l)}</div>{note}</div>')
    out.append("</div>")
    return "".join(out)


def _table(df, tid, money_cols=(), pct_cols=(), max_rows=500):
    if df is None or len(df) == 0:
        return '<p class="muted">No data.</p>'
    df = df.head(max_rows)
    head = "".join(f"<th onclick=\"srt('{tid}',{i})\">{_e(c)}</th>" for i, c in enumerate(df.columns))
    body = []
    for _, r in df.iterrows():
        cells = []
        for c in df.columns:
            v = r[c]
            if c in pct_cols:
                try:
                    cells.append(f'<td class="n" data-v="{float(v)}">{_p(100 * float(v))}</td>')
                except (TypeError, ValueError):
                    cells.append('<td class="n">-</td>')
            elif c in money_cols:
                try:
                    cells.append(f'<td class="n" data-v="{float(v)}">{_m(v)}</td>')
                except (TypeError, ValueError):
                    cells.append('<td class="n">-</td>')
            else:
                cells.append(f"<td>{_e(v)}</td>")
        body.append("<tr>" + "".join(cells) + "</tr>")
    return (f'<input class="search" placeholder="Search this table" oninput="flt(\'{tid}\',this.value)">'
            f'<div class="tw"><table id="{tid}"><thead><tr>{head}</tr></thead><tbody>{"".join(body)}</tbody></table></div>')


def _heat(df, label_col, month_cols, top=10):
    if not month_cols:
        return ""
    mx = float(df[month_cols].max().max()) or 1.0
    head = "<th>" + _e(label_col) + "</th>" + "".join(f"<th>{_e(c)}</th>" for c in month_cols)
    rows = []
    for _, r in df.head(top).iterrows():
        tds = "".join(f'<td class="n" style="background:rgba(11,127,138,{min(0.85, 0.08 + 0.77 * float(r[c]) / mx):.2f})">{_m(r[c])}</td>' for c in month_cols)
        rows.append(f"<tr><td>{_e(r[label_col])}</td>{tds}</tr>")
    return f'<div class="tw"><table><thead><tr>{head}</tr></thead><tbody>{"".join(rows)}</tbody></table></div>'


def _fy_table(c):
    svc, today = c["svc"], c["today"]
    monthly = svc.groupby("Month")["Cost"].sum()
    fm = c["full_months"]
    daily = None
    if len(fm) >= 3:
        start = pd.Timestamp(fm[-3])
        days = (pd.Timestamp(c["cur_month"]) - start).days
        daily = float(monthly.reindex(fm[-3:]).sum()) / days if days > 0 else None
    g = c["growth"] / 100.0
    y0 = today.year
    rows, prev = [], None
    for y in range(y0 - 2, y0 + 3):
        basis, actual, fc, total = "", None, None, None
        if y in c["fy_history"]:
            actual = total = c["fy_history"][y]; fc = 0.0; basis = "Entered from invoices"
        elif y < y0:
            months_y = [m for m in monthly.index if pd.Timestamp(m).year == y]
            if len(months_y) == 12:
                actual = total = float(monthly.reindex(months_y).sum()); fc = 0.0; basis = "Actual"
            else:
                basis = "No data (older than the cost history; use --fy-history)"
        elif y == y0:
            ytd = float(monthly[[m for m in monthly.index if pd.Timestamp(m).year == y]].sum())
            rem = (pd.Timestamp(y, 12, 31) - pd.Timestamp(today)).days + 1
            actual = ytd
            if daily is not None:
                fc = rem * daily * (1 + g) ** ((rem / 2) / 365.0)
                total = ytd + fc
            basis = "Actual to date + forecast"
        else:
            if daily is not None:
                total = fc = 365 * daily * (1 + g) ** (y - y0)
            basis = f"Forecast (last-3-month run-rate, growth {c['growth']:g}% a year)"
        chg = None if (total is None or prev is None or prev <= 0) else (total - prev) / prev
        rows.append({"Fiscal year": f"FY{y}", "Basis": basis, "Actual to date": actual, "Forecast (remaining)": fc, "Total": total,
                     "Change vs previous year": (None if (total is None or prev is None) else total - prev), "Change %": chg})
        if total is not None:
            prev = total
    return pd.DataFrame(rows), daily


def _tag_view(df):
    if df is None or len(df) == 0 or "Cost" not in df.columns:
        return None
    col = "TagValue" if "TagValue" in df.columns else df.columns[1]
    d = df.copy()
    d["Value"] = d[col].fillna("").astype(str).str.strip().str.lower()
    d.loc[d["Value"] == "", "Value"] = "(no value)"
    return d.groupby("Value")["Cost"].sum().sort_values(ascending=False)


CSS = """
:root{--navy:#0b1f3a;--teal:#0b7f8a;--amber:#f2a33a;--bg:#f4f7fa;--mut:#5b6b7f}
*{box-sizing:border-box}body{margin:0;font-family:Segoe UI,Calibri,Arial,sans-serif;background:var(--bg);color:#16253a}
header{background:var(--navy);color:#fff;padding:22px 32px}header h1{margin:0 0 4px;font-size:26px}header p{margin:0;color:#c9d6e6;font-size:14px}
nav{position:sticky;top:0;background:#fff;border-bottom:1px solid #d8e0ea;padding:8px 24px;z-index:5;overflow-x:auto;white-space:nowrap}
nav a{margin-right:18px;color:var(--teal);text-decoration:none;font-weight:600;font-size:14px}
main{max-width:1280px;margin:0 auto;padding:20px 24px 60px}
section{background:#fff;border-radius:8px;padding:18px 22px;margin:18px 0;box-shadow:0 1px 3px rgba(0,0,0,.06)}
h2{margin:0 0 4px;font-size:20px;color:var(--navy)}h3{margin:18px 0 6px;font-size:15px;color:var(--navy)}
.muted{color:var(--mut);font-size:13px}.note{background:#fff7e8;border-left:4px solid var(--amber);padding:8px 12px;margin:8px 0;font-size:13px}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:12px;margin-top:10px}
.kpi{background:var(--bg);border-radius:8px;padding:12px 14px}.kl{font-size:12px;color:var(--mut)}.kv{font-size:26px;font-weight:700;color:var(--navy);margin:2px 0}.ks{font-size:12px;color:var(--mut)}
.two{display:grid;grid-template-columns:1fr 1fr;gap:24px}@media(max-width:900px){.two{grid-template-columns:1fr}}
.bars .bar{display:flex;align-items:center;gap:8px;margin:4px 0}.lbl{width:240px;font-size:13px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.track{flex:1;background:#e8eef5;border-radius:4px;height:16px}.fill{height:16px;background:var(--teal);border-radius:4px}
.val{width:190px;text-align:right;font-size:13px}.sh{color:var(--mut);margin-left:4px}
.cols{display:flex;align-items:flex-end;gap:8px;height:300px;padding-top:20px}.col{flex:1;display:flex;flex-direction:column;justify-content:flex-end;align-items:center;height:100%;min-width:36px}
.cb{width:70%;background:var(--teal);border-radius:4px 4px 0 0;min-height:1px}.cv{font-size:11px;margin-bottom:2px}.cl{font-size:11px;margin-top:4px}.cn{font-size:11px;color:var(--mut);height:14px}
.tw{overflow:auto;max-height:520px;margin-top:6px}table{border-collapse:collapse;width:100%;font-size:13px}
th{position:sticky;top:0;background:var(--navy);color:#fff;text-align:left;padding:6px 8px;cursor:pointer;white-space:nowrap}
td{padding:5px 8px;border-bottom:1px solid #e6ecf3}td.n{text-align:right;font-variant-numeric:tabular-nums}tr:hover td{background:#f0f6fa}
.search{width:260px;padding:6px 8px;border:1px solid #c9d4e0;border-radius:6px;margin-top:6px}
"""
JS = """
function flt(id,q){q=q.toLowerCase();document.querySelectorAll('#'+id+' tbody tr').forEach(function(r){r.style.display=r.innerText.toLowerCase().indexOf(q)>-1?'':'none'})}
function srt(id,i){var t=document.getElementById(id),b=t.tBodies[0],rows=Array.prototype.slice.call(b.rows),asc=t.getAttribute('data-s')!=i+'a';
rows.sort(function(x,y){var a=x.cells[i],c=y.cells[i],av=a.getAttribute('data-v'),cv=c.getAttribute('data-v');
if(av!==null&&cv!==null){return asc?av-cv:cv-av}return asc?a.innerText.localeCompare(c.innerText):c.innerText.localeCompare(a.innerText)});
rows.forEach(function(r){b.appendChild(r)});t.setAttribute('data-s',i+(asc?'a':'d'))}
"""


def write_html(path, c):
    today, cur = c["today"], pd.Timestamp(c["cur_month"])
    ccy = c["currency"] or ""
    svc = c["svc"]
    n_full, last12, lf, prev_t = c["n_full"], c["last12"], c["lf_total"], c["prev_total"]
    mom = None if not prev_t else 100 * (lf - prev_t) / prev_t
    elapsed = max(today.day - 1, 1)
    dim = (pd.Timestamp(today) + pd.offsets.MonthEnd(0)).day
    eom = c["mtd"] / elapsed * dim if c["mtd"] else None
    subs, err = c["subs"], c["errors_df"]
    with_cost = svc["SubscriptionId"].nunique()
    failed = err["SubscriptionId"].nunique() if len(err) else 0
    adv_total = c["adv_total"]
    fy_df, daily = _fy_table(c)

    kpis = "".join([
        _kpi(f"Spend, last {n_full} full months", _m(last12), ccy),
        _kpi(f"Last full month ({pd.Timestamp(c['last_full']):%b %Y})", _m(lf), ("Month on month " + _p(mom)) if mom is not None else ""),
        _kpi("Month to date", _m(c["mtd"]), f"{cur:%b %Y}, {elapsed} day(s) of data"),
        _kpi("End-of-month forecast", _m(eom), "run-rate estimate"),
        _kpi("Advisor potential savings (annual)", _m(adv_total), (f"{100 * adv_total / last12:.1f}% of last-12-month spend" if (last12 and adv_total and n_full == 12) else "identified, not realized")),
        _kpi("Subscriptions", f"{len(subs)}", f"{with_cost} with cost data, {failed} with failed queries"),
    ])
    parts = [f'<section id="summary"><h2>Executive summary</h2><p class="muted">Cost view: {_e(c["cost_type"])}. Currency: {_e(ccy)}. '
             f'Generated {today:%d %b %Y}. Savings are potential (Advisor) until verified in later billing data.</p><div class="kpis">{kpis}</div></section>']

    bs = c["by_service"].reset_index()
    first = bs.columns[0]
    mcols = [x for x in bs.columns if isinstance(x, str) and len(x) == 7 and x[4] == "-"]
    full_cols = [x for x in mcols if pd.Timestamp(x + "-01") < cur][-12:]
    bs["Period total"] = bs[full_cols].sum(axis=1) if full_cols else bs["Total"]
    bs = bs.sort_values("Period total", ascending=False)
    tot_all = float(bs["Period total"].sum())
    top15 = [(r[first], r["Period total"]) for _, r in bs.head(15).iterrows()]
    other = tot_all - sum(v for _, v in top15)
    parts.append(f'<section id="drivers"><h2>Top 15 cost drivers by Azure service</h2><p class="muted">Last {len(full_cols) or "all"} full months. Share is of total spend across all services.</p>'
                 + _hbars(top15, total=tot_all) + (f'<p class="muted">All other services: {_m(other)} ({100 * other / tot_all:.1f}%)</p>' if tot_all and other > 0 else "") + "</section>")

    mt = svc.groupby("Month")["Cost"].sum().sort_index()
    labels = [f"{pd.Timestamp(m):%b %y}" + ("*" if pd.Timestamp(m) >= cur else "") for m in mt.index]
    vals = list(mt.values)
    notes = [""] + [_p(100 * (vals[i] - vals[i - 1]) / vals[i - 1]) if (vals[i - 1] and pd.Timestamp(mt.index[i]) < cur) else "" for i in range(1, len(vals))]
    parts.append('<section id="trend"><h2>Monthly trend and month-on-month change</h2><p class="muted">* current month is partial.</p>'
                 + _columns(labels, vals, notes) + "<h3>Top 10 services by month</h3>"
                 + _heat(bs.rename(columns={first: "Service"}), "Service", mcols[-13:], 10) + "</section>")

    note = ("Forecast = run-rate of the last 3 full months" + (f" ({_m(daily * 365)} a year)" if daily else "") +
            f", growth {c['growth']:g}% a year. Years older than the cost history need <code>--fy-history</code> (invoice totals), for example <code>--fy-history \"2024=1200000,2025=1500000\"</code>.")
    parts.append('<section id="forecast"><h2>Forecast and year-on-year</h2><div class="note">' + note + "</div>"
                 + _table(fy_df, "t_fy", money_cols=("Actual to date", "Forecast (remaining)", "Total", "Change vs previous year"), pct_cols=("Change %",)) + "</section>")

    sub_tbl = c["by_sub"].reset_index(); sub_tbl = sub_tbl.rename(columns={sub_tbl.columns[0]: "Subscription"})
    sub_top = [(r["Subscription"], r["Total"]) for _, r in sub_tbl.head(15).iterrows()]
    keep = ["Subscription", "Total", "Share %"] + (["MoM % (last full month)"] if "MoM % (last full month)" in sub_tbl.columns else [])
    offer_tbl = c["by_offer"].reset_index(); offer_tbl = offer_tbl.rename(columns={offer_tbl.columns[0]: "Agreement type"})[["Agreement type", "Total", "Share %"]]
    env_tbl = c["by_env"].reset_index(); env_tbl = env_tbl.rename(columns={env_tbl.columns[0]: "Environment"})[["Environment", "Total", "Share %"]]
    parts.append('<section id="subs"><h2>Subscriptions</h2><div class="two"><div><h3>Top 15 by spend</h3>' + _hbars(sub_top) + "</div><div><h3>By agreement type</h3>"
                 + _table(offer_tbl, "t_offer", money_cols=("Total",), pct_cols=("Share %",)) + "<h3>Dev/Test vs other offers</h3>"
                 + _table(env_tbl, "t_env", money_cols=("Total",), pct_cols=("Share %",)) + "</div></div><h3>All subscriptions</h3>"
                 + _table(sub_tbl[keep], "t_sub", money_cols=("Total",), pct_cols=tuple(x for x in keep if "%" in x)) + "</section>")

    rg = c["by_rg"].reset_index(); rg = rg.rename(columns={rg.columns[0]: "Resource group"})
    tr = c["top_res"]; res_html = ""
    if len(tr):
        trr = tr.copy(); trr["Resource"] = trr["ResourceId"].astype(str).str.split("/").str[-1]
        trr["Resource group"] = trr["ResourceId"].astype(str).str.extract(r"(?i)/resourcegroups/([^/]+)")[0]
        res_html = ("<h3>Top 20 resources (last full month)</h3>" + _table(trr[["SubscriptionName", "Resource group", "Resource", "ServiceName", "Cost"]].head(20), "t_res20", money_cols=("Cost",))
                    + "<h3>Top 300 resources</h3>" + _table(trr[["SubscriptionName", "Resource group", "Resource", "ServiceName", "Cost", "ResourceId"]], "t_res", money_cols=("Cost",)))
    parts.append('<section id="rgs"><h2>Resource groups and resources</h2><h3>Top 15 resource groups</h3>' + _hbars([(r["Resource group"], r["Total"]) for _, r in rg.head(15).iterrows()]) + res_html + "</section>")

    pr = c["pricing"]; rate_html = '<p class="muted">No pricing-model data.</p>'
    if len(pr) and "PricingModel" in pr.columns:
        lfp = pr[pr["Month"] == pd.Timestamp(c["last_full"])].assign(PricingModel=lambda d: d["PricingModel"].fillna("Unknown")).groupby("PricingModel")["Cost"].sum().sort_values(ascending=False)
        rate_html = "<h3>Spend by pricing model (last full month)</h3>" + _hbars(list(lfp.items()))
        if len(c["cov"]):
            rate_html += "<h3>Commitment coverage by month</h3>" + _table(c["cov"], "t_cov", money_cols=tuple(x for x in c["cov"].columns if x not in ("Month", "Commitment-covered %")), pct_cols=("Commitment-covered %",))
    adv = c["adv"]
    if len(adv):
        ca = adv[adv["solution"].fillna("").str.lower().str.contains("reserv|savings plan", regex=True)]
        if len(ca):
            ca = ca.assign(Type=ca["solution"].str.lower().str.contains("savings plan").map({True: "Savings plans", False: "Reservations"}))
            rate_html += "<h3>Advisor commitment recommendations</h3>" + _table(ca.groupby("Type")["annualSavings"].agg(["count", "sum"]).rename(columns={"count": "Recommendations", "sum": "Annual savings (potential)"}).reset_index(), "t_ca", money_cols=("Annual savings (potential)",))
    parts.append('<section id="rate"><h2>Rate optimization: Reservations and Savings Plans</h2>' + rate_html + "</section>")

    tag_html = ""
    for k, df in c["tags"].items():
        tv = _tag_view(df)
        if tv is None:
            tag_html += f"<h3>{_e(k)}</h3><p class='muted'>No data for this tag key.</p>"; continue
        tag_html += f"<h3>Cost by {_e(k)} (last full month; values trimmed and lower-cased so duplicates merge)</h3>" + _hbars(list(tv.head(12).items()), total=float(tv.sum()))
    parts.append('<section id="tags"><h2>Tag analytics</h2>' + (tag_html or '<p class="muted">Tags were not collected. Run with <code>--tag-keys Owner,Application,Team</code>.</p>') + "</section>")

    rg_df = c["region"]
    if len(rg_df) and "ResourceLocation" in rg_df.columns:
        rr = rg_df.assign(Region=rg_df["ResourceLocation"].fillna("(none)")).groupby("Region")["Cost"].sum().sort_values(ascending=False)
        parts.append('<section id="regions"><h2>Region analytics</h2>' + _hbars(list(rr.head(12).items()), total=float(rr.sum())) + "</section>")

    an = c["anomalies"]
    parts.append('<section id="anomalies"><h2>Cost anomalies</h2><p class="muted">Service and subscription pairs where the last full month is more than 30% and 500 above the average of the 3 months before.</p>'
                 + (_table(an, "t_an", money_cols=("Last full month", "Avg prior 3 months", "Change $"), pct_cols=("Change %",)) if len(an) else "<p class='muted'>No anomalies found (or fewer than 4 full months of data).</p>") + "</section>")

    sav = ""
    if len(c["adv_svc"]):
        a = c["adv_svc"].reset_index()
        sav += "<h3>Potential annual savings by service</h3>" + _hbars([(r.iloc[0], r.iloc[1]) for _, r in a.head(12).iterrows()], total=float(a.iloc[:, 1].sum()))
        sav += "<h3>Top recommendations</h3>" + _table(c["adv_top"].head(50), "t_adv", money_cols=("annualSavings",))
        if len(c["savings_vs_spend"]):
            sv = c["savings_vs_spend"]
            sav += "<h3>Savings compared with spend (approximate service-name match)</h3>" + _table(sv, "t_sv", money_cols=tuple(sv.columns[1:3]), pct_cols=("Savings as % of spend",))
    parts.append('<section id="savings"><h2>Savings opportunities (Advisor)</h2><div class="note">Stages are not additive: potential, approved, implemented and realized are the same savings at different stages.</div>' + (sav or "<p class='muted'>No Advisor data.</p>") + "</section>")

    cm = c["commit_map"]
    if len(cm):
        cols = [x for x in ["Subscription", "Recommendation", "Term", "Region", "SKU", "Match quality", "Resource group", "Resource name", "Resource SKU", "Last-month cost", "Cost share %", "Allocated annual savings (estimate)"] if x in cm.columns]
        parts.append('<section id="commit"><h2>Reservation and savings plan recommendations mapped to resources</h2><div class="note">Azure does not attribute commitments to resources. Rows list the resources that make up the usage; the savings split is an estimate by cost share.</div>'
                     + _table(cm[cols], "t_cm", money_cols=("Last-month cost", "Allocated annual savings (estimate)"), pct_cols=("Cost share %",), max_rows=1500) + "</section>")

    have = set(svc["SubscriptionId"].unique())
    miss = subs[~subs["SubscriptionId"].isin(have)][["SubscriptionName", "State", "Offer", "SubscriptionId"]]
    cov_html = (f"<p>{len(subs)} subscriptions were listed; <b>{with_cost}</b> returned cost data; <b>{len(miss)}</b> did not.</p>"
                + ("<h3>Subscriptions without cost data</h3>" + _table(miss, "t_miss") if len(miss) else "")
                + ("<h3>Errors (reasons)</h3>" + _table(err, "t_err") if len(err) else ""))
    parts.append('<section id="coverage"><h2>Data coverage and quality</h2>' + cov_html + "</section>")

    nav = "".join(f'<a href="#{i}">{n}</a>' for i, n in [("summary", "Summary"), ("drivers", "Top 15"), ("trend", "Trend"), ("forecast", "Forecast"), ("subs", "Subscriptions"),
          ("rgs", "Resources"), ("rate", "Rate optimization"), ("tags", "Tags"), ("regions", "Regions"), ("anomalies", "Anomalies"), ("savings", "Savings"), ("commit", "Commitments"), ("coverage", "Coverage")])
    doc = (f'<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Azure FinOps Dashboard</title><style>{CSS}</style></head><body>'
           f'<header><h1>Azure FinOps Dashboard</h1><p>Spend, savings and forecast across {len(subs)} subscriptions. Generated {today:%d %b %Y}.</p></header><nav>{nav}</nav><main>{"".join(parts)}</main><script>{JS}</script></body></html>')
    with open(path, "w", encoding="utf-8") as f:
        f.write(doc)


def main():
    ap = argparse.ArgumentParser(description="Azure FinOps collector and report")
    ap.add_argument("--out", default="finops_output")
    ap.add_argument("--months", type=int, default=13, help="months incl. current (Azure keeps ~13)")
    ap.add_argument("--cost-type", default="ActualCost", choices=["ActualCost", "AmortizedCost"])
    ap.add_argument("--max-subs", type=int, default=0, help="limit subscriptions (for a quick test)")
    ap.add_argument("--pause", type=float, default=1.0, help="seconds between queries")
    ap.add_argument("--report-only", action="store_true", help="skip Azure calls; use cache")
    ap.add_argument("--tag-keys", default="Owner,Application,Team", help="comma list of tag keys for cost-by-tag (empty to skip)")
    ap.add_argument("--growth", type=float, default=0.0, help="yearly growth %% used in the forecast (default 0)")
    ap.add_argument("--fy-history", default="", help='known past-year totals from invoices, e.g. "2024=1200000,2025=1500000"')
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
    tag_keys = [k.strip() for k in args.tag_keys.split(",") if k.strip()]
    fy_hist = {}
    for part in args.fy_history.split(","):
        if "=" in part:
            y, v = part.split("=", 1)
            try:
                fy_hist[int(y.strip())] = float(v.strip().replace(" ", ""))
            except ValueError:
                pass

    errors = []
    if not args.report_only:
        errors = collect_cost(c, work, args.out, args.months, args.cost_type, today, args.pause, tag_keys)
        collect_advisor(c, list(work["SubscriptionId"]), args.out)
        collect_inventory(c, list(work["SubscriptionId"]), args.out)
        err_f = os.path.join(args.out, "errors.csv")
        old = pd.read_csv(err_f) if os.path.exists(err_f) else pd.DataFrame()
        errs = pd.concat([old, pd.DataFrame(errors)], ignore_index=True).drop_duplicates() if errors or len(old) else pd.DataFrame()
        errs.to_csv(err_f, index=False)
    err_f = os.path.join(args.out, "errors.csv")
    errors_df = pd.read_csv(err_f) if os.path.exists(err_f) and os.path.getsize(err_f) > 2 else pd.DataFrame()

    build_report(args.out, args.cost_type, work, errors_df, today, tag_keys, args.growth, fy_hist)


if __name__ == "__main__":
    main()

"""
Dynamics 365 CRM MCP Server
Connects to Microsoft Dynamics 365 CRM via OData API with MSAL auth.
Provides tools for querying accounts, opportunities, contacts, milestones,
deal teams, and pipeline data.
"""

import os
import logging
from typing import Optional

import msal
import requests
from fastmcp import FastMCP

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

BASE_URL = "https://microsoftsales.crm.dynamics.com"
API_URL = f"{BASE_URL}/api/data/v9.2"
CLIENT_ID = "51f81489-12ee-4a9e-aaae-a2591f45987d"
TENANT_ID = "72f988bf-86f1-41af-91ab-2d7cd011db47"
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = [f"{BASE_URL}/user_impersonation"]
USER_ID = "d8229040-69cf-f011-bbd3-7c1e5257b8e3"

CACHE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".token_cache.json")

logger = logging.getLogger("dynamics-mcp")

# ---------------------------------------------------------------------------
# Dynamics 365 Client
# ---------------------------------------------------------------------------


class DynamicsClient:
    """Handles MSAL authentication and OData API calls to Dynamics 365."""

    def __init__(self):
        self._cache = msal.SerializableTokenCache()
        if os.path.exists(CACHE_FILE):
            with open(CACHE_FILE, "r") as f:
                self._cache.deserialize(f.read())
        self._app = msal.PublicClientApplication(
            CLIENT_ID, authority=AUTHORITY, token_cache=self._cache,
        )
        self._token: Optional[str] = None

    # -- auth ---------------------------------------------------------------

    def _save_cache(self):
        if self._cache.has_state_changed:
            with open(CACHE_FILE, "w") as f:
                f.write(self._cache.serialize())

    def _ensure_token(self):
        """Acquire a token silently (cached/refresh) or fall back to interactive."""
        accounts = self._app.get_accounts()
        result = (
            self._app.acquire_token_silent(SCOPES, account=accounts[0])
            if accounts
            else None
        )
        if not result:
            result = self._app.acquire_token_interactive(SCOPES, prompt="select_account")
        if "access_token" not in result:
            raise RuntimeError(f"Auth failed: {result.get('error_description', result)}")
        self._save_cache()
        self._token = result["access_token"]

    @property
    def headers(self) -> dict:
        self._ensure_token()
        return {
            "Authorization": f"Bearer {self._token}",
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0",
            "Accept": "application/json",
            "Prefer": "odata.include-annotations=*,odata.maxpagesize=500",
        }

    # -- generic OData helpers ----------------------------------------------

    def get(self, entity: str, params: Optional[dict] = None) -> list[dict]:
        """GET with automatic pagination. Returns all records."""
        url = f"{API_URL}/{entity}"
        all_records: list[dict] = []
        while url:
            resp = requests.get(url, headers=self.headers, params=params)
            resp.raise_for_status()
            data = resp.json()
            all_records.extend(data.get("value", []))
            url = data.get("@odata.nextLink")
            params = None  # nextLink already contains params
        return all_records

    def get_with_count(self, entity: str, params: Optional[dict] = None) -> tuple[list[dict], int]:
        """GET first page and return (records, total_count). Use $count=true in params."""
        url = f"{API_URL}/{entity}"
        resp = requests.get(url, headers=self.headers, params=params)
        resp.raise_for_status()
        data = resp.json()
        count = data.get("@odata.count", len(data.get("value", [])))
        records = data.get("value", [])
        # follow pages
        next_link = data.get("@odata.nextLink")
        while next_link:
            resp = requests.get(next_link, headers=self.headers)
            resp.raise_for_status()
            data = resp.json()
            records.extend(data.get("value", []))
            next_link = data.get("@odata.nextLink")
        return records, count

    # -- record cleaning ----------------------------------------------------

    @staticmethod
    def _formatted_key(key: str) -> str:
        return f"{key}@OData.Community.Display.V1.FormattedValue"

    @staticmethod
    def clean_record(raw: dict, fields: list[str]) -> dict:
        """Extract *fields* from *raw*, preferring OData formatted values."""
        out: dict = {}
        for f in fields:
            fmt_key = f"{f}@OData.Community.Display.V1.FormattedValue"
            if fmt_key in raw:
                out[f] = raw[fmt_key]
            elif f in raw:
                out[f] = raw[f]
        return out

    # -- entity-specific queries -------------------------------------------

    def my_account_ids(self) -> list[dict]:
        """Return msp_accountteams rows for the current user.
        Falls back to deriving accounts from deal team memberships if
        msp_accountteams is inaccessible."""
        try:
            filt = f"_msp_systemuserid_value eq '{USER_ID}'"
            rows = self.get("msp_accountteams", {"$filter": filt})
            if rows:
                return rows
        except Exception:
            pass
        # Fallback: return empty — callers should use my_account_tree() instead
        return []

    def my_account_tree(self) -> dict:
        """Derive the user's account tree from deal team memberships.

        Returns dict with:
            parent_accounts: list of top-level parent account dicts
            all_account_ids: set of ALL account IDs (parents + children)
            parent_ids: set of just parent account IDs
            child_map: dict mapping child_id -> parent_id
            name_map: dict mapping account_id -> account name
        """
        # Step 1: Get my deal team entries -> opportunity IDs
        filt = f"_msp_dealteamuserid_value eq '{USER_ID}'"
        deal_rows = self.get("msp_dealteams", {"$filter": filt})
        opp_ids = list({
            r["_msp_parentopportunityid_value"]
            for r in deal_rows
            if r.get("_msp_parentopportunityid_value")
        })
        if not opp_ids:
            return {
                "parent_accounts": [], "all_account_ids": set(),
                "parent_ids": set(), "child_map": {}, "name_map": {},
            }

        # Step 2: Get accounts from those opportunities
        opp_clauses = " or ".join(f"opportunityid eq '{oid}'" for oid in opp_ids)
        opps = self.get("opportunities", {
            "$filter": opp_clauses,
            "$select": "opportunityid,_parentaccountid_value",
        })
        acct_ids = list({
            o["_parentaccountid_value"] for o in opps
            if o.get("_parentaccountid_value")
        })
        if not acct_ids:
            return {
                "parent_accounts": [], "all_account_ids": set(),
                "parent_ids": set(), "child_map": {}, "name_map": {},
            }

        # Step 3: Fetch those accounts to find their parents
        acct_clauses = " or ".join(f"accountid eq '{a}'" for a in acct_ids)
        select = (
            "accountid,name,accountnumber,msp_parentinglevelcode,"
            "msp_endcustomersegmentcode,msp_industrycode,msp_managedstatuscode,"
            "openrevenue,_parentaccountid_value,msp_activecontacts,"
            "address1_city,address1_country,statecode"
        )
        accts = self.get("accounts", {"$filter": acct_clauses, "$select": select})

        # Step 4: Traverse up to top-level parents
        parent_ids: set[str] = set()
        seen_accts = {a["accountid"]: a for a in accts}

        for acct in accts:
            parent_ref = acct.get("_parentaccountid_value")
            if parent_ref and parent_ref not in seen_accts:
                # Fetch the parent
                try:
                    parent_acct = self.get_single(f"accounts({parent_ref})")
                    if parent_acct:
                        seen_accts[parent_acct["accountid"]] = parent_acct
                        parent_ids.add(parent_acct["accountid"])
                except Exception:
                    parent_ids.add(acct["accountid"])
            elif not parent_ref:
                # This account IS a top-level parent
                parent_ids.add(acct["accountid"])
            else:
                parent_ids.add(parent_ref)

        # Step 5: Get all children under each parent
        all_ids: set[str] = set(parent_ids)
        child_map: dict[str, str] = {}
        name_map: dict[str, str] = {}

        for pid in parent_ids:
            children = self.get_child_accounts(pid)
            for c in children:
                cid = c["accountid"]
                all_ids.add(cid)
                child_map[cid] = pid
                name_map[cid] = c.get("name", cid)

        # Add parent names
        for pid in parent_ids:
            if pid in seen_accts:
                name_map[pid] = seen_accts[pid].get("name", pid)
            else:
                name_map[pid] = pid

        # Build parent account list
        parent_accounts = []
        for pid in parent_ids:
            if pid in seen_accts:
                parent_accounts.append(seen_accts[pid])
            else:
                parent_accounts.append({"accountid": pid, "name": name_map.get(pid, pid)})

        return {
            "parent_accounts": parent_accounts,
            "all_account_ids": all_ids,
            "parent_ids": parent_ids,
            "child_map": child_map,
            "name_map": name_map,
        }

    def search_account(self, name: str) -> list[dict]:
        """Case-insensitive contains search on account name."""
        filt = f"contains(name, '{name}')"
        select = (
            "accountid,name,accountnumber,msp_parentinglevelcode,"
            "msp_endcustomersegmentcode,msp_industrycode,msp_managedstatuscode,"
            "openrevenue,_parentaccountid_value,msp_activecontacts,"
            "address1_city,address1_country,telephone1,fax,statecode"
        )
        return self.get("accounts", {"$filter": filt, "$select": select})

    def get_child_accounts(self, parent_id: str) -> list[dict]:
        """Get accounts whose parent is *parent_id*."""
        filt = f"_parentaccountid_value eq '{parent_id}'"
        select = "accountid,name,accountnumber,msp_endcustomersegmentcode,msp_industrycode,openrevenue,statecode"
        return self.get("accounts", {"$filter": filt, "$select": select})

    def get_opportunities(
        self,
        account_ids: list[str],
        statecode: Optional[int] = None,
    ) -> list[dict]:
        """Get opportunities for one or more account IDs with optional status filter."""
        if not account_ids:
            return []
        # Build account filter with OR
        acct_clauses = " or ".join(
            f"_parentaccountid_value eq '{aid}'" for aid in account_ids
        )
        filt = f"({acct_clauses})"
        if statecode is not None:
            filt += f" and statecode eq {statecode}"
        select = (
            "opportunityid,name,description,_parentaccountid_value,"
            "estimatedvalue,estimatedclosedate,actualvalue,actualclosedate,"
            "msp_activesalesstage,msp_activeprocess,msp_billedrevenue,"
            "msp_billedrevenuestatus,statecode,statuscode"
        )
        return self.get("opportunities", {
            "$filter": filt,
            "$select": select,
            "$orderby": "estimatedclosedate asc",
        })

    def get_contacts(self, account_id: str) -> list[dict]:
        filt = f"_parentcustomerid_value eq '{account_id}'"
        select = "contactid,fullname,firstname,lastname,emailaddress1,telephone1,jobtitle,msp_jobrolecode"
        return self.get("contacts", {"$filter": filt, "$select": select})

    def get_account_team(self, account_id: str) -> list[dict]:
        try:
            filt = f"_msp_accountid_value eq '{account_id}'"
            return self.get("msp_accountteams", {"$filter": filt})
        except Exception:
            return []

    # -- single record & write ops -----------------------------------------

    def get_single(self, entity: str, record_id: str, params: Optional[dict] = None) -> dict:
        """GET a single record by ID."""
        url = f"{API_URL}/{entity}({record_id})"
        resp = requests.get(url, headers=self.headers, params=params)
        resp.raise_for_status()
        return resp.json()

    def post(self, entity: str, data: dict) -> dict:
        """POST to create a record."""
        url = f"{API_URL}/{entity}"
        headers = {**self.headers, "Content-Type": "application/json"}
        resp = requests.post(url, headers=headers, json=data)
        resp.raise_for_status()
        if resp.status_code == 204:
            return {"status": "created", "entity_id": resp.headers.get("OData-EntityId", "")}
        return resp.json() if resp.content else {"status": "created"}

    def patch(self, entity: str, record_id: str, data: dict) -> dict:
        """PATCH to update a record."""
        url = f"{API_URL}/{entity}({record_id})"
        headers = {**self.headers, "Content-Type": "application/json"}
        resp = requests.patch(url, headers=headers, json=data)
        resp.raise_for_status()
        return {"status": "updated", "record_id": record_id}

    # -- metadata discovery ------------------------------------------------

    def get_entity_definitions(self, search: Optional[str] = None) -> list[dict]:
        """Query EntityDefinitions to discover available entities.
        Fetches all and filters client-side (metadata API has limited $filter support).
        """
        url = f"{API_URL}/EntityDefinitions"
        params: dict = {
            "$select": "LogicalName,DisplayName,Description,EntitySetName,IsCustomEntity",
        }
        resp = requests.get(url, headers=self.headers, params=params)
        resp.raise_for_status()
        data = resp.json()
        results = []
        for e in data.get("value", []):
            logical = e.get("LogicalName", "")
            display = e.get("DisplayName", {}).get("UserLocalizedLabel") or {}
            desc = e.get("Description", {}).get("UserLocalizedLabel") or {}
            display_label = display.get("Label") or ""
            desc_label = desc.get("Label") or ""
            if search:
                term = search.lower()
                if not (term in logical.lower()
                        or term in display_label.lower()
                        or term in desc_label.lower()):
                    continue
            results.append({
                "logical_name": logical,
                "display_name": display_label or None,
                "description": desc_label or None,
                "entity_set_name": e.get("EntitySetName"),
                "is_custom": e.get("IsCustomEntity"),
            })
        return results

    def get_entity_attributes(self, entity_logical_name: str) -> list[dict]:
        """Query attribute metadata for a specific entity."""
        url = f"{API_URL}/EntityDefinitions(LogicalName='{entity_logical_name}')/Attributes"
        params = {
            "$select": "LogicalName,DisplayName,AttributeType,Description,IsCustomAttribute",
        }
        resp = requests.get(url, headers=self.headers, params=params)
        resp.raise_for_status()
        data = resp.json()
        results = []
        for a in data.get("value", []):
            display = a.get("DisplayName", {}).get("UserLocalizedLabel") or {}
            desc = a.get("Description", {}).get("UserLocalizedLabel") or {}
            results.append({
                "logical_name": a.get("LogicalName"),
                "display_name": display.get("Label"),
                "attribute_type": a.get("AttributeType"),
                "description": desc.get("Label"),
                "is_custom": a.get("IsCustomAttribute"),
            })
        return results

    # -- relationships & hierarchy -----------------------------------------

    def get_connections(self, record_id: str) -> list[dict]:
        """Get connections for a record (deal teams, relationships, etc.)."""
        filt = f"_record1id_value eq '{record_id}' or _record2id_value eq '{record_id}'"
        return self.get("connections", {"$filter": filt})

    def get_parent_account(self, account_id: str) -> Optional[dict]:
        """Get the parent account of an account, if one exists."""
        try:
            acct = self.get_single("accounts", account_id,
                                   {"$select": "_parentaccountid_value"})
            parent_id = acct.get("_parentaccountid_value")
            if parent_id:
                select = (
                    "accountid,name,accountnumber,msp_parentinglevelcode,"
                    "msp_endcustomersegmentcode,msp_industrycode,msp_managedstatuscode,"
                    "openrevenue,_parentaccountid_value,msp_activecontacts,"
                    "address1_city,address1_country,telephone1,fax,statecode"
                )
                return self.get_single("accounts", parent_id, {"$select": select})
        except Exception:
            pass
        return None

    def search_opportunities_by_name(self, name: str) -> list[dict]:
        """Search opportunities by name (case-insensitive contains)."""
        filt = f"contains(name, '{name}')"
        select = (
            "opportunityid,name,description,_parentaccountid_value,"
            "estimatedvalue,estimatedclosedate,actualvalue,actualclosedate,"
            "msp_activesalesstage,msp_activeprocess,msp_billedrevenue,"
            "msp_billedrevenuestatus,statecode,statuscode"
        )
        return self.get("opportunities", {"$filter": filt, "$select": select})


# ---------------------------------------------------------------------------
# Singleton client
# ---------------------------------------------------------------------------

client = DynamicsClient()

# ---------------------------------------------------------------------------
# FastMCP Server
# ---------------------------------------------------------------------------

mcp = FastMCP("Dynamics 365 CRM")

ACCOUNT_FIELDS = [
    "accountid", "name", "accountnumber", "msp_parentinglevelcode",
    "msp_endcustomersegmentcode", "msp_industrycode", "msp_managedstatuscode",
    "openrevenue", "_parentaccountid_value", "msp_activecontacts",
    "address1_city", "address1_country", "telephone1", "fax", "statecode",
]

OPP_FIELDS = [
    "opportunityid", "name", "description", "_parentaccountid_value",
    "estimatedvalue", "estimatedclosedate", "actualvalue", "actualclosedate",
    "msp_activesalesstage", "msp_activeprocess", "msp_billedrevenue",
    "msp_billedrevenuestatus", "statecode", "statuscode",
]

CONTACT_FIELDS = [
    "contactid", "fullname", "firstname", "lastname",
    "emailaddress1", "telephone1", "jobtitle", "msp_jobrolecode",
]

TEAM_FIELDS = [
    "_msp_accountid_value", "_msp_systemuserid_value",
    "msp_rolename", "msp_roletype", "msp_solutionarea",
    "msp_fullname", "msp_title",
]

CONNECTION_FIELDS = [
    "connectionid", "_record1id_value", "_record2id_value",
    "_record1roleid_value", "_record2roleid_value", "description",
]


def _resolve_account(name: str) -> dict:
    """Find a single account by name. Raises if not found or ambiguous."""
    matches = client.search_account(name)
    if not matches:
        raise ValueError(f"No account found matching '{name}'")
    # Prefer exact (case-insensitive) match, else return first
    for m in matches:
        if m.get("name", "").lower() == name.lower():
            return m
    return matches[0]


STATUS_MAP = {"open": 0, "won": 1, "lost": 2, "all": None}


def _resolve_opportunity(name: str) -> dict:
    """Find a single opportunity by name. Raises if not found."""
    matches = client.search_opportunities_by_name(name)
    if not matches:
        raise ValueError(f"No opportunity found matching '{name}'")
    for m in matches:
        if m.get("name", "").lower() == name.lower():
            return m
    return matches[0]


# ---------------------------------------------------------------------------
# Tools
# ---------------------------------------------------------------------------


@mcp.tool
def my_accounts() -> list[dict]:
    """List all Dynamics 365 accounts assigned to me (Alexander Shaul) via msp_accountteams.

    Returns each account's name, segment, industry, managed status,
    open revenue, and parenting level.
    """
    tree = client.my_account_tree()
    parent_accounts = tree["parent_accounts"]
    if not parent_accounts:
        return []

    results = []
    for acct in parent_accounts:
        cleaned = client.clean_record(acct, ACCOUNT_FIELDS)
        # Fetch child accounts for this parent
        children = client.get_child_accounts(acct["accountid"])
        cleaned["child_account_count"] = len(children)
        cleaned["child_accounts"] = [
            {"accountid": c["accountid"], "name": c.get("name")}
            for c in children
        ]
        results.append(cleaned)

    return results


@mcp.tool
def account_details(account_name: str) -> dict:
    """Search for a Dynamics 365 account by name (case-insensitive partial match)
    and return full details including parent account hierarchy and child accounts.

    Args:
        account_name: Full or partial account name to search for.
    """
    acct = _resolve_account(account_name)
    cleaned = client.clean_record(acct, ACCOUNT_FIELDS)

    # Fetch parent account info
    parent_raw = client.get_parent_account(acct["accountid"])
    if parent_raw:
        cleaned["parent_account"] = client.clean_record(parent_raw, ACCOUNT_FIELDS)
    else:
        cleaned["parent_account"] = None

    # Fetch child accounts
    children_raw = client.get_child_accounts(acct["accountid"])
    child_fields = ["accountid", "name", "accountnumber", "msp_endcustomersegmentcode",
                     "msp_industrycode", "openrevenue", "statecode"]
    cleaned["child_accounts"] = [client.clean_record(c, child_fields) for c in children_raw]
    cleaned["child_account_count"] = len(children_raw)
    cleaned["is_parent"] = len(children_raw) > 0
    cleaned["is_child"] = parent_raw is not None

    return cleaned


@mcp.tool
def account_opportunities(
    account_name: str,
    include_children: bool = True,
    status: str = "open",
) -> dict:
    """Get opportunities for a Dynamics 365 account.

    Args:
        account_name: Full or partial account name.
        include_children: If True, also include opportunities from child accounts.
        status: Filter by status — "open", "won", "lost", or "all".
    """
    acct = _resolve_account(account_name)
    acct_id = acct["accountid"]
    acct_ids = [acct_id]

    children = []
    if include_children:
        children = client.get_child_accounts(acct_id)
        acct_ids.extend(c["accountid"] for c in children)

    statecode = STATUS_MAP.get(status.lower())
    raw_opps = client.get_opportunities(acct_ids, statecode=statecode)
    opps = [client.clean_record(o, OPP_FIELDS) for o in raw_opps]

    return {
        "account": acct.get("name"),
        "account_id": acct_id,
        "status_filter": status,
        "include_children": include_children,
        "child_accounts_searched": len(children),
        "opportunity_count": len(opps),
        "opportunities": opps,
    }


@mcp.tool
def search_opportunities(
    query: Optional[str] = None,
    min_value: Optional[float] = None,
    stage: Optional[str] = None,
    closing_before: Optional[str] = None,
) -> dict:
    """Search opportunities across all my assigned accounts with optional filters.

    Args:
        query: Text to search in opportunity names (case-insensitive).
        min_value: Minimum estimated value to include.
        stage: Filter by sales stage (e.g. "Listen & Consult", "Inspire & Design").
        closing_before: ISO date string (e.g. "2025-12-31") — only opps closing before this date.
    """
    tree = client.my_account_tree()
    all_ids = tree["all_account_ids"]
    if not all_ids:
        return {"opportunity_count": 0, "opportunities": []}

    # Build OData filter
    id_clauses = " or ".join(f"_parentaccountid_value eq '{a}'" for a in all_ids)
    filters = [f"({id_clauses})", "statecode eq 0"]

    if query:
        filters.append(f"contains(name, '{query}')")
    if min_value is not None:
        filters.append(f"estimatedvalue ge {min_value}")
    if closing_before:
        filters.append(f"estimatedclosedate le {closing_before}")

    filt = " and ".join(filters)
    select = (
        "opportunityid,name,description,_parentaccountid_value,"
        "estimatedvalue,estimatedclosedate,msp_activesalesstage,"
        "msp_activeprocess,statecode,statuscode"
    )
    raw = client.get("opportunities", {
        "$filter": filt,
        "$select": select,
        "$orderby": "estimatedvalue desc",
    })

    opps = [client.clean_record(o, OPP_FIELDS) for o in raw]

    # Post-filter by stage (formatted value) since OData may not support contains on optionset
    if stage:
        stage_lower = stage.lower()
        opps = [o for o in opps if stage_lower in str(o.get("msp_activesalesstage", "")).lower()]

    return {
        "filters_applied": {
            "query": query,
            "min_value": min_value,
            "stage": stage,
            "closing_before": closing_before,
        },
        "opportunity_count": len(opps),
        "opportunities": opps,
    }


@mcp.tool
def pipeline_summary() -> dict:
    """Aggregate pipeline summary across all my assigned accounts.

    Shows per-account totals, per-stage breakdown, and grand total
    of open opportunity estimated values.
    """
    tree = client.my_account_tree()
    parent_ids = tree["parent_ids"]
    all_ids = tree["all_account_ids"]
    child_map = tree["child_map"]
    name_map = tree["name_map"]

    if not all_ids:
        return {"total_pipeline": 0, "by_account": {}, "by_stage": {}}

    # Fetch open opps
    raw_opps = client.get_opportunities(list(all_ids), statecode=0)

    by_account: dict[str, float] = {}
    by_stage: dict[str, float] = {}
    grand_total = 0.0

    for opp in raw_opps:
        val = opp.get("estimatedvalue") or 0
        grand_total += val

        # Map opp to parent account name
        opp_acct_id = opp.get("_parentaccountid_value")
        parent_id = child_map.get(opp_acct_id, opp_acct_id)
        acct_name = name_map.get(parent_id, opp.get(
            "_parentaccountid_value@OData.Community.Display.V1.FormattedValue", parent_id or "Unknown"
        ))
        by_account[acct_name] = by_account.get(acct_name, 0) + val

        # Stage
        stage = opp.get(
            "msp_activesalesstage@OData.Community.Display.V1.FormattedValue",
            opp.get("msp_activesalesstage", "Unknown"),
        )
        by_stage[stage] = by_stage.get(stage, 0) + val

    # Sort by value descending
    by_account = dict(sorted(by_account.items(), key=lambda x: x[1], reverse=True))
    by_stage = dict(sorted(by_stage.items(), key=lambda x: x[1], reverse=True))

    return {
        "total_pipeline": grand_total,
        "account_count": len(by_account),
        "opportunity_count": len(raw_opps),
        "by_account": by_account,
        "by_stage": by_stage,
    }


@mcp.tool
def account_contacts(account_name: str) -> dict:
    """Get all contacts for a Dynamics 365 account.

    Args:
        account_name: Full or partial account name to search for.
    """
    acct = _resolve_account(account_name)
    raw = client.get_contacts(acct["accountid"])
    contacts = [client.clean_record(c, CONTACT_FIELDS) for c in raw]

    return {
        "account": acct.get("name"),
        "account_id": acct["accountid"],
        "contact_count": len(contacts),
        "contacts": contacts,
    }


@mcp.tool
def account_team(account_name: str) -> dict:
    """Get the full account team for a Dynamics 365 account from msp_accountteams.

    Shows all team members, their roles, solution areas, and titles.
    If the account has no direct team, automatically checks the parent account.
    If msp_accountteams is inaccessible, returns an appropriate message.

    Args:
        account_name: Full or partial account name to search for.
    """
    acct = _resolve_account(account_name)
    raw = client.get_account_team(acct["accountid"])
    members = [client.clean_record(m, TEAM_FIELDS) for m in raw]

    result = {
        "account": acct.get("name"),
        "account_id": acct["accountid"],
        "team_member_count": len(members),
        "team_members": members,
        "inherited_from_parent": False,
    }

    # If no team found, traverse up to parent account
    if not members:
        parent = client.get_parent_account(acct["accountid"])
        if parent:
            parent_raw = client.get_account_team(parent["accountid"])
            parent_members = [client.clean_record(m, TEAM_FIELDS) for m in parent_raw]
            if parent_members:
                result["team_members"] = parent_members
                result["team_member_count"] = len(parent_members)
                result["inherited_from_parent"] = True
                result["parent_account"] = parent.get("name")
                result["parent_account_id"] = parent.get("accountid")

    if not result["team_members"]:
        result["note"] = "No team members found. The msp_accountteams entity may be inaccessible."

    return result


# ---------------------------------------------------------------------------
# Discovery & Exploration Tools
# ---------------------------------------------------------------------------


@mcp.tool
def discover_entities(search_term: str) -> dict:
    """Search for Dynamics 365 entities by logical name.

    Use this to find entities for milestones, deal teams, activities, etc.

    Args:
        search_term: Text to search in entity logical names (e.g. "milestone", "team", "opportunity").
    """
    results = client.get_entity_definitions(search=search_term)
    return {
        "search_term": search_term,
        "entity_count": len(results),
        "entities": results,
    }


@mcp.tool
def discover_fields(entity_name: str, search_term: Optional[str] = None) -> dict:
    """Get fields/attributes for a Dynamics 365 entity.

    Args:
        entity_name: The logical name of the entity (e.g. "opportunity", "msp_milestone").
        search_term: Optional text to filter field names (case-insensitive).
    """
    results = client.get_entity_attributes(entity_name)
    if search_term:
        term = search_term.lower()
        results = [
            r for r in results
            if term in (r.get("logical_name") or "").lower()
            or term in (r.get("display_name") or "").lower()
        ]
    return {
        "entity_name": entity_name,
        "field_count": len(results),
        "fields": results,
    }


@mcp.tool
def run_odata_query(
    entity_set: str,
    odata_filter: Optional[str] = None,
    select: Optional[str] = None,
    top: Optional[int] = 50,
    orderby: Optional[str] = None,
    expand: Optional[str] = None,
) -> dict:
    """Run a custom OData query against Dynamics 365.

    Use this for ad-hoc queries when no specific tool exists.
    Combine with discover_entities and discover_fields to explore the data model.

    Args:
        entity_set: The entity set name (e.g. "opportunities", "connections", "msp_milestones").
        odata_filter: OData $filter expression.
        select: Comma-separated field names for $select.
        top: Maximum records to return (default 50).
        orderby: OData $orderby expression.
        expand: OData $expand expression for related entities.
    """
    params: dict = {}
    if odata_filter:
        params["$filter"] = odata_filter
    if select:
        params["$select"] = select
    if top:
        params["$top"] = str(top)
    if orderby:
        params["$orderby"] = orderby
    if expand:
        params["$expand"] = expand

    try:
        records = client.get(entity_set, params if params else None)
        return {
            "entity_set": entity_set,
            "record_count": len(records),
            "records": records,
        }
    except requests.HTTPError as e:
        return {
            "entity_set": entity_set,
            "error": str(e),
            "hint": "Use discover_entities to find valid entity set names.",
        }


# ---------------------------------------------------------------------------
# Opportunity Tools
# ---------------------------------------------------------------------------


@mcp.tool
def opportunity_detail(opportunity_name: str) -> dict:
    """Get full details for a single opportunity including connections and relationships.

    Args:
        opportunity_name: Full or partial opportunity name to search for.
    """
    opp = _resolve_opportunity(opportunity_name)
    cleaned = client.clean_record(opp, OPP_FIELDS)

    # Get deal team from msp_dealteams
    try:
        deal_team_raw = client.get("msp_dealteams", {
            "$filter": f"_msp_parentopportunityid_value eq '{opp['opportunityid']}'",
        })
        team = []
        for dt in deal_team_raw:
            team.append({
                "user_id": dt.get("_msp_dealteamuserid_value"),
                "user_name": dt.get(
                    "_msp_dealteamuserid_value@OData.Community.Display.V1.FormattedValue",
                    dt.get("_msp_dealteamuserid_value"),
                ),
                "date_added": dt.get(
                    "msp_dateadded@OData.Community.Display.V1.FormattedValue",
                    dt.get("msp_dateadded"),
                ),
                "is_owner": dt.get(
                    "msp_isowner@OData.Community.Display.V1.FormattedValue",
                    dt.get("msp_isowner"),
                ),
                "is_me": dt.get("_msp_dealteamuserid_value") == USER_ID,
            })
        cleaned["deal_team"] = team
        cleaned["deal_team_count"] = len(team)
        cleaned["i_am_on_team"] = any(m.get("is_me") for m in team)
    except Exception as e:
        cleaned["deal_team"] = []
        cleaned["deal_team_count"] = 0
        cleaned["deal_team_error"] = str(e)

    return cleaned


@mcp.tool
def opportunity_team(opportunity_name: str) -> dict:
    """Get deal team members for an opportunity via msp_dealteams.

    Shows all people on the deal team, their roles,
    and whether the current user (Alexander Shaul) is on the team.

    Args:
        opportunity_name: Full or partial opportunity name.
    """
    opp = _resolve_opportunity(opportunity_name)
    opp_id = opp["opportunityid"]

    try:
        deal_team_raw = client.get("msp_dealteams", {
            "$filter": f"_msp_parentopportunityid_value eq '{opp_id}'",
        })
    except Exception:
        deal_team_raw = []

    team = []
    for dt in deal_team_raw:
        team.append({
            "user_id": dt.get("_msp_dealteamuserid_value"),
            "user_name": dt.get(
                "_msp_dealteamuserid_value@OData.Community.Display.V1.FormattedValue",
                dt.get("_msp_dealteamuserid_value"),
            ),
            "date_added": dt.get(
                "msp_dateadded@OData.Community.Display.V1.FormattedValue",
                dt.get("msp_dateadded"),
            ),
            "is_owner": dt.get(
                "msp_isowner@OData.Community.Display.V1.FormattedValue",
                dt.get("msp_isowner"),
            ),
            "is_me": dt.get("_msp_dealteamuserid_value") == USER_ID,
        })

    return {
        "opportunity": opp.get("name"),
        "opportunity_id": opp_id,
        "team_member_count": len(team),
        "i_am_on_team": any(m.get("is_me") for m in team),
        "team_members": team,
    }


@mcp.tool
def my_opportunities() -> dict:
    """List all opportunities where I (Alexander Shaul) am connected as a deal team member.

    Searches msp_dealteams where my user ID appears, then resolves linked opportunities.
    """
    filt = f"_msp_dealteamuserid_value eq '{USER_ID}'"
    try:
        deal_rows = client.get("msp_dealteams", {"$filter": filt})
    except Exception as e:
        return {"error": str(e), "opportunities": []}

    opp_ids = list({
        r["_msp_parentopportunityid_value"]
        for r in deal_rows
        if r.get("_msp_parentopportunityid_value")
    })

    if not opp_ids:
        return {"opportunity_count": 0, "opportunities": []}

    id_clauses = " or ".join(f"opportunityid eq '{oid}'" for oid in opp_ids)
    select = (
        "opportunityid,name,description,_parentaccountid_value,"
        "estimatedvalue,estimatedclosedate,msp_activesalesstage,"
        "msp_activeprocess,statecode,statuscode"
    )
    try:
        raw_opps = client.get("opportunities", {"$filter": id_clauses, "$select": select})
    except Exception:
        raw_opps = []

    opps = [client.clean_record(o, OPP_FIELDS) for o in raw_opps]

    return {
        "opportunity_count": len(opps),
        "opportunities": opps,
    }


@mcp.tool
def opportunities_not_on_team() -> dict:
    """List open opportunities across my assigned accounts where I am NOT on the deal team.

    Compares all open opportunities under my accounts (including child accounts)
    against my msp_dealteams entries to find gaps.
    """
    tree = client.my_account_tree()
    all_ids = tree["all_account_ids"]
    if not all_ids:
        return {"opportunity_count": 0, "opportunities": []}

    # Get all open opportunities for these accounts
    all_opps = client.get_opportunities(list(all_ids), statecode=0)

    # Get my deal team entries
    filt = f"_msp_dealteamuserid_value eq '{USER_ID}'"
    try:
        deal_rows = client.get("msp_dealteams", {"$filter": filt})
    except Exception:
        deal_rows = []

    my_opp_ids = {
        r["_msp_parentopportunityid_value"]
        for r in deal_rows
        if r.get("_msp_parentopportunityid_value")
    }

    # Filter to opps I'm NOT on the deal team for
    not_on_team = [
        o for o in all_opps
        if o.get("opportunityid") not in my_opp_ids
    ]
    opps = [client.clean_record(o, OPP_FIELDS) for o in not_on_team]

    return {
        "total_account_opportunities": len(all_opps),
        "on_deal_team": len(all_opps) - len(not_on_team),
        "not_on_deal_team": len(not_on_team),
        "opportunities": opps,
    }


@mcp.tool
def opportunity_milestones(opportunity_name: str) -> dict:
    """Get milestones for an opportunity.

    Auto-discovers milestone entities in Dynamics 365 and queries them
    for the given opportunity. Also falls back to checking activities/tasks.

    Args:
        opportunity_name: Full or partial opportunity name.
    """
    opp = _resolve_opportunity(opportunity_name)
    opp_id = opp["opportunityid"]

    MILESTONE_SELECT = (
        "msp_engagementmilestoneid,msp_name,msp_milestonestatus,"
        "msp_milestonedate,msp_milestonecategory,msp_workload,"
        "msp_monthlyuse,msp_milestonesolutionarea,msp_commitmentrecommendation,"
        "msp_forecastcomments,msp_milestonecomments,"
        "_ownerid_value,_msp_opportunityid_value,_msp_parentaccount_value,"
        "statecode,statuscode"
    )

    try:
        raw = client.get("msp_engagementmilestones", {
            "$filter": f"_msp_opportunityid_value eq '{opp_id}'",
            "$select": MILESTONE_SELECT,
        })
    except Exception as e:
        return {
            "opportunity": opp.get("name"),
            "opportunity_id": opp_id,
            "error": str(e),
            "milestones": [],
        }

    MILESTONE_FIELDS = [
        "msp_engagementmilestoneid", "msp_name", "msp_milestonestatus",
        "msp_milestonedate", "msp_milestonecategory", "msp_workload",
        "msp_monthlyuse", "msp_milestonesolutionarea", "msp_commitmentrecommendation",
        "msp_forecastcomments", "msp_milestonecomments",
        "_ownerid_value", "_msp_opportunityid_value", "_msp_parentaccount_value",
        "statecode", "statuscode",
    ]
    milestones = [client.clean_record(m, MILESTONE_FIELDS) for m in raw]

    return {
        "opportunity": opp.get("name"),
        "opportunity_id": opp_id,
        "milestone_count": len(milestones),
        "milestones": milestones,
    }


@mcp.tool
def my_milestones() -> dict:
    """Get all milestones across opportunities where I (Alexander Shaul) am on the deal team.

    Finds my opportunities via msp_dealteams, then fetches milestones for each one.
    Returns milestones grouped by opportunity.
    """
    # Step 1: Find my deal team entries
    filt = f"_msp_dealteamuserid_value eq '{USER_ID}'"
    try:
        deal_rows = client.get("msp_dealteams", {"$filter": filt})
    except Exception as e:
        return {"error": f"Failed to query deal teams: {e}", "opportunities": []}

    opp_ids = list({
        r["_msp_parentopportunityid_value"]
        for r in deal_rows
        if r.get("_msp_parentopportunityid_value")
    })

    if not opp_ids:
        return {"opportunity_count": 0, "total_milestone_count": 0, "opportunities": []}

    # Step 2: Resolve opportunity names
    id_clauses = " or ".join(f"opportunityid eq '{oid}'" for oid in opp_ids)
    select = (
        "opportunityid,name,_parentaccountid_value,"
        "estimatedvalue,msp_activesalesstage,statecode"
    )
    try:
        raw_opps = client.get("opportunities", {"$filter": id_clauses, "$select": select})
    except Exception:
        raw_opps = []

    # Step 3: Fetch milestones for all opportunities in one query
    ms_clauses = " or ".join(f"_msp_opportunityid_value eq '{oid}'" for oid in opp_ids)
    ms_select = (
        "msp_engagementmilestoneid,msp_name,msp_milestonestatus,"
        "msp_milestonedate,msp_milestonecategory,msp_workload,"
        "msp_monthlyuse,msp_milestonesolutionarea,msp_commitmentrecommendation,"
        "msp_forecastcomments,_ownerid_value,_msp_opportunityid_value,"
        "statecode,statuscode"
    )
    try:
        all_ms_raw = client.get("msp_engagementmilestones", {
            "$filter": ms_clauses,
            "$select": ms_select,
        })
    except Exception:
        all_ms_raw = []

    # Group milestones by opportunity
    ms_by_opp: dict[str, list[dict]] = {}
    MILESTONE_FIELDS = [
        "msp_engagementmilestoneid", "msp_name", "msp_milestonestatus",
        "msp_milestonedate", "msp_milestonecategory", "msp_workload",
        "msp_monthlyuse", "msp_milestonesolutionarea", "msp_commitmentrecommendation",
        "msp_forecastcomments", "_ownerid_value", "_msp_opportunityid_value",
        "statecode", "statuscode",
    ]
    for ms in all_ms_raw:
        opp_id = ms.get("_msp_opportunityid_value")
        cleaned = client.clean_record(ms, MILESTONE_FIELDS)
        ms_by_opp.setdefault(opp_id, []).append(cleaned)

    # Step 4: Build results
    results = []
    total_milestones = 0
    for opp in raw_opps:
        opp_id = opp["opportunityid"]
        opp_name = opp.get(
            "name@OData.Community.Display.V1.FormattedValue",
            opp.get("name", opp_id),
        )
        opp_stage = opp.get(
            "msp_activesalesstage@OData.Community.Display.V1.FormattedValue",
            opp.get("msp_activesalesstage"),
        )
        opp_account = opp.get(
            "_parentaccountid_value@OData.Community.Display.V1.FormattedValue",
            opp.get("_parentaccountid_value"),
        )

        milestones = ms_by_opp.get(opp_id, [])
        total_milestones += len(milestones)
        results.append({
            "opportunity": opp_name,
            "opportunity_id": opp_id,
            "account": opp_account,
            "stage": opp_stage,
            "milestone_count": len(milestones),
            "milestones": milestones,
        })

    return {
        "opportunity_count": len(raw_opps),
        "total_milestone_count": total_milestones,
        "opportunities": results,
    }


# ---------------------------------------------------------------------------
# Contact & Account Lookup
# ---------------------------------------------------------------------------


@mcp.tool
def find_contact_by_email(email: str) -> dict:
    """Find a Dynamics 365 contact by email address and return their parent account.

    Searches across emailaddress1, emailaddress2, and emailaddress3 fields.
    Returns contact details and the linked parent account if available.

    Args:
        email: Email address to search for (case-insensitive).
    """
    email_lower = email.strip().lower()
    filt = (
        f"contains(emailaddress1, '{email_lower}') or "
        f"contains(emailaddress2, '{email_lower}') or "
        f"contains(emailaddress3, '{email_lower}')"
    )
    select = (
        "contactid,fullname,emailaddress1,emailaddress2,emailaddress3,"
        "jobtitle,_parentcustomerid_value"
    )
    try:
        contacts = client.get("contacts", {"$filter": filt, "$select": select})
    except requests.HTTPError as e:
        return {"error": str(e), "hint": "Check email format and permissions."}

    if not contacts:
        return {"found": False, "email": email, "contacts": []}

    results = []
    for c in contacts:
        entry = client.clean_record(c, [
            "contactid", "fullname", "emailaddress1", "emailaddress2",
            "emailaddress3", "jobtitle", "_parentcustomerid_value",
        ])
        # Resolve parent account name
        acct_id = c.get("_parentcustomerid_value")
        fmt_key = "_parentcustomerid_value@OData.Community.Display.V1.FormattedValue"
        entry["parent_account_name"] = c.get(fmt_key, "")
        entry["parent_account_id"] = acct_id or ""
        results.append(entry)

    return {"found": True, "email": email, "contact_count": len(results), "contacts": results}


@mcp.tool
def find_account_by_domain(domain: str) -> dict:
    """Find Dynamics 365 accounts whose website or email domains match the given domain.

    Useful for matching meeting attendees to accounts when you only have
    an email domain (e.g. 'contoso.com').

    Args:
        domain: Domain to search for (e.g. 'contoso.com').
    """
    domain_lower = domain.strip().lower()
    filt = (
        f"contains(websiteurl, '{domain_lower}') or "
        f"contains(emailaddress1, '{domain_lower}')"
    )
    select = "accountid,name,websiteurl,emailaddress1,_parentaccountid_value"
    try:
        accounts = client.get("accounts", {"$filter": filt, "$select": select})
    except requests.HTTPError as e:
        return {"error": str(e), "hint": "Check domain format and permissions."}

    if not accounts:
        return {"found": False, "domain": domain, "accounts": []}

    results = []
    for a in accounts:
        entry = client.clean_record(a, [
            "accountid", "name", "websiteurl", "emailaddress1",
            "_parentaccountid_value",
        ])
        fmt_key = "_parentaccountid_value@OData.Community.Display.V1.FormattedValue"
        entry["parent_account_name"] = a.get(fmt_key, "")
        results.append(entry)

    return {"found": True, "domain": domain, "account_count": len(results), "accounts": results}


# ---------------------------------------------------------------------------
# Annotation / Notes Operations
# ---------------------------------------------------------------------------


@mcp.tool
def create_note(
    regarding_entity: str,
    regarding_id: str,
    note_text: str,
    subject: str = "",
) -> dict:
    """Create an annotation (note) on a Dynamics 365 record.

    Use this to add comments/notes to milestones, opportunities, or any entity
    that supports annotations. The note will appear in the record's timeline.

    Args:
        regarding_entity: The logical entity name the note is about
            (e.g. 'msp_milestone', 'opportunity', 'account').
        regarding_id: The GUID of the record the note relates to.
        note_text: The body text of the note (supports plain text).
        subject: Optional subject/title for the note.
    """
    payload = {
        f"objectid_{regarding_entity}@odata.bind": f"/{regarding_entity}s({regarding_id})",
        "notetext": note_text,
    }
    if subject:
        payload["subject"] = subject

    try:
        result = client.post("annotations", payload)
        return {
            "status": "created",
            "regarding_entity": regarding_entity,
            "regarding_id": regarding_id,
            "subject": subject,
            "entity_id": result.get("entity_id", ""),
        }
    except requests.HTTPError as e:
        error_body = ""
        if hasattr(e, "response") and e.response is not None:
            try:
                error_body = e.response.json()
            except Exception:
                error_body = e.response.text[:500]
        return {
            "error": str(e),
            "error_detail": error_body,
            "hint": (
                "Check that the entity supports annotations and that the "
                "regarding_entity uses the correct logical name (singular, "
                "e.g. 'msp_milestone' not 'msp_milestones'). "
                "Use discover_entities to verify."
            ),
        }


@mcp.tool
def search_annotations(
    regarding_entity: str,
    regarding_id: str,
    search_text: str = "",
    top: int = 20,
) -> dict:
    """Search annotations (notes) on a Dynamics 365 record.

    Use this to check for existing notes before creating duplicates,
    or to review the note history on a milestone/opportunity.

    Args:
        regarding_entity: The logical entity name (e.g. 'msp_milestone', 'opportunity').
        regarding_id: The GUID of the record to search notes for.
        search_text: Optional text to filter notes by (searches subject and notetext).
        top: Maximum number of notes to return (default 20).
    """
    filt = f"_objectid_value eq '{regarding_id}'"
    if search_text:
        filt += f" and (contains(subject, '{search_text}') or contains(notetext, '{search_text}'))"

    params: dict = {
        "$filter": filt,
        "$select": "annotationid,subject,notetext,createdon,_createdby_value",
        "$orderby": "createdon desc",
        "$top": str(top),
    }
    try:
        notes = client.get("annotations", params)
    except requests.HTTPError as e:
        return {"error": str(e), "hint": "Check entity and record ID."}

    results = []
    for n in notes:
        entry = client.clean_record(n, [
            "annotationid", "subject", "notetext", "createdon", "_createdby_value",
        ])
        fmt_key = "_createdby_value@OData.Community.Display.V1.FormattedValue"
        entry["created_by_name"] = n.get(fmt_key, "")
        results.append(entry)

    return {
        "regarding_entity": regarding_entity,
        "regarding_id": regarding_id,
        "note_count": len(results),
        "notes": results,
    }


@mcp.tool
def create_record(entity_set: str, fields: str) -> dict:
    """Create a new record in Dynamics 365.

    Use discover_entities and discover_fields to find the correct entity set
    and field names before creating.

    Args:
        entity_set: The entity set name (e.g. 'annotations', 'tasks', 'phonecalls').
        fields: JSON string of field names and values.
            Example: '{"subject": "Follow-up", "description": "Call notes..."}'
            For lookup fields, use OData bind syntax:
            '{"objectid_msp_milestone@odata.bind": "/msp_milestones(GUID)"}'
    """
    import json as _json

    try:
        field_data = _json.loads(fields)
    except _json.JSONDecodeError as e:
        return {"error": f"Invalid JSON in fields: {e}", "hint": "Provide a valid JSON string."}

    if not isinstance(field_data, dict):
        return {"error": "fields must be a JSON object (dict), not a list or scalar."}

    try:
        result = client.post(entity_set, field_data)
        return {
            "status": "created",
            "entity_set": entity_set,
            "entity_id": result.get("entity_id", ""),
        }
    except requests.HTTPError as e:
        error_body = ""
        if hasattr(e, "response") and e.response is not None:
            try:
                error_body = e.response.json()
            except Exception:
                error_body = e.response.text[:500]
        return {
            "error": str(e),
            "error_detail": error_body,
            "hint": "Use discover_fields to check valid field names for this entity.",
        }


# ---------------------------------------------------------------------------
# Write Operations
# ---------------------------------------------------------------------------


@mcp.tool
def update_record(entity_set: str, record_id: str, field_name: str, field_value: str) -> dict:
    """Update a single field on a Dynamics 365 record.

    Use discover_fields to find the correct field names before updating.

    Args:
        entity_set: The entity set name (e.g. "msp_milestones", "tasks").
        record_id: The GUID of the record to update.
        field_name: The field to update (e.g. "subject", "description").
        field_value: The new value for the field.
    """
    try:
        client.patch(entity_set, record_id, {field_name: field_value})
        return {
            "entity_set": entity_set,
            "record_id": record_id,
            "status": "updated",
            "field_updated": field_name,
        }
    except requests.HTTPError as e:
        return {
            "entity_set": entity_set,
            "record_id": record_id,
            "error": str(e),
            "hint": "Check field names with discover_fields and ensure you have write permission.",
        }


@mcp.tool
def assign_to_me(entity_set: str, record_id: str) -> dict:
    """Assign a Dynamics 365 record (milestone, task, etc.) to me (Alexander Shaul).

    Sets the ownerid field to my user ID.

    Args:
        entity_set: The entity set name (e.g. "msp_milestones", "tasks").
        record_id: The GUID of the record to assign to me.
    """
    updates = {
        "ownerid@odata.bind": f"/systemusers({USER_ID})",
    }
    try:
        client.patch(entity_set, record_id, updates)
        return {
            "entity_set": entity_set,
            "record_id": record_id,
            "status": "assigned_to_me",
            "user_id": USER_ID,
        }
    except requests.HTTPError as e:
        return {
            "entity_set": entity_set,
            "record_id": record_id,
            "error": str(e),
            "hint": "The entity may use a different assignment field. Use discover_fields to check.",
        }


# ---------------------------------------------------------------------------
# Entrypoint
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    mcp.run()

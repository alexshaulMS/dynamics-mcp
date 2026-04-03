# Dynamics 365 CRM MCP Server

An MCP (Model Context Protocol) server that connects to Microsoft Dynamics 365 CRM (MSX) via the OData API. Designed for use with GitHub Copilot CLI and other MCP-compatible AI assistants.

## Features

### Account Management
- **`my_accounts`** — List all parent accounts derived from your deal team memberships, with child accounts
- **`account_details`** — Full account details including parent/child hierarchy
- **`account_contacts`** — Get contacts for an account
- **`account_team`** — Account team members (auto-traverses to parent if child has no team)
- **`account_opportunities`** — Opportunities for an account (including child accounts)

### Opportunities & Pipeline
- **`my_opportunities`** — Opportunities where you're on the deal team
- **`opportunities_not_on_team`** — Open opportunities in your accounts that you're NOT on
- **`opportunity_detail`** — Full opportunity detail with deal team
- **`opportunity_team`** — Deal team members for an opportunity
- **`search_opportunities`** — Search with filters (query, stage, value, close date)
- **`pipeline_summary`** — Aggregate pipeline by account and stage

### Milestones
- **`my_milestones`** — All milestones across your deal team opportunities
- **`opportunity_milestones`** — Milestones for a specific opportunity

### Discovery & Exploration
- **`discover_entities`** — Search Dynamics metadata for entity names
- **`discover_fields`** — Get fields/attributes for any entity
- **`run_odata_query`** — Run custom OData queries for ad-hoc exploration

### Write Operations
- **`update_record`** — Update a field on any record
- **`assign_to_me`** — Assign a record (milestone, task, etc.) to yourself

## Prerequisites

- Python 3.10+
- [GitHub Copilot CLI](https://docs.github.com/en/copilot/github-copilot-in-the-cli) installed
- Access to a Dynamics 365 CRM instance (e.g. `microsoftsales.crm.dynamics.com`)
- An Azure AD app registration client ID with `Dynamics CRM user_impersonation` delegated permission (for Microsoft internal users, ask your team for the shared public client app ID)

## Setup

### 1. Clone and install dependencies

```bash
git clone https://github.com/alexshaulMS/dynamics-mcp.git
cd dynamics-mcp
pip install -r requirements.txt
```

### 2. Find your Dynamics User ID

You need your `systemuserid` GUID. Open this URL in your browser while logged into Dynamics:

```
https://microsoftsales.crm.dynamics.com/api/data/v9.2/systemusers?$filter=internalemailaddress eq 'your.email@microsoft.com'&$select=systemuserid,fullname
```

Copy the `systemuserid` value from the response.

### 3. Add to GitHub Copilot CLI

Add the following to your `~/.copilot/mcp-config.json` (create the file if it doesn't exist):

```json
{
  "mcpServers": {
    "dynamics-crm": {
      "command": "python",
      "args": ["C:\\full\\path\\to\\dynamics-mcp\\server.py"],
      "env": {
        "DYNAMICS_CLIENT_ID": "your-azure-ad-app-client-id",
        "DYNAMICS_TENANT_ID": "your-azure-ad-tenant-id",
        "DYNAMICS_USER_ID": "your-systemuserid-guid"
      },
      "tools": ["*"]
    }
  }
}
```

> **Note:** Use the full absolute path to `server.py`. On Windows use double backslashes.

### 4. Restart and authenticate

Restart GitHub Copilot CLI. On first launch, a browser window will open for Azure AD login. After authenticating, your token is cached in `.token_cache.json` and refreshed automatically — you won't need to log in again.

### 5. Verify it works

In Copilot CLI, try:
```
show me my accounts
```
or
```
what opportunities am I on the deal team for?
```

## Configuration Reference

| Environment Variable | Description | Required |
|---------------------|-------------|----------|
| `DYNAMICS_CLIENT_ID` | Azure AD app registration client ID | **Yes** |
| `DYNAMICS_TENANT_ID` | Azure AD tenant ID | **Yes** |
| `DYNAMICS_USER_ID` | Your Dynamics `systemuserid` GUID | **Yes** |
| `DYNAMICS_BASE_URL` | Dynamics instance URL | No (defaults to `https://microsoftsales.crm.dynamics.com`) |

## Architecture

```
┌─────────────────┐     stdio      ┌──────────────────┐     OData v9.2     ┌───────────────┐
│  Copilot CLI /   │◄──────────────►│  FastMCP Server  │◄──────────────────►│  Dynamics 365  │
│  MCP Client      │                │  (server.py)     │     MSAL auth      │  CRM (MSX)     │
└─────────────────┘                └──────────────────┘                    └───────────────┘
```

## Entity Reference

| Entity | Entity Set | Used For |
|--------|-----------|----------|
| `msp_dealteam` | `msp_dealteams` | Opportunity deal team members |
| `msp_engagementmilestone` | `msp_engagementmilestones` | Opportunity milestones |
| `msp_accountteam` | `msp_accountteams` | Account team assignments |
| `opportunity` | `opportunities` | Sales opportunities |
| `account` | `accounts` | Customer accounts |
| `contact` | `contacts` | Account contacts |

## Troubleshooting

- **"DYNAMICS_CLIENT_ID environment variable is required"** — You haven't set the env vars. Add them to the `env` block in your `mcp-config.json`.
- **Browser doesn't open for login** — Make sure you're running the server in an environment that can open a browser. The first auth requires interactive login.
- **Token expired** — Delete `.token_cache.json` and restart. It will prompt for login again.
- **"Not connected" in Copilot CLI** — Restart Copilot CLI. MCP server reconnection sometimes requires a full restart.

## License

Internal use only.

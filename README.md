# Dynamics 365 CRM MCP Server

An MCP (Model Context Protocol) server that connects to Microsoft Dynamics 365 CRM (MSX) via the OData API. Designed for use with GitHub Copilot CLI and other MCP-compatible AI assistants.

## Features

### Account Management
- **`my_accounts`** — List all accounts assigned to you via `msp_accountteams`
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
- Access to a Dynamics 365 CRM instance
- An Azure AD app registration with `user_impersonation` permissions on your Dynamics instance

## Setup

### 1. Clone and install dependencies

```bash
git clone <this-repo>
cd dynamics-mcp
pip install -r requirements.txt
```

### 2. Create a `.env` file

Copy `.env.example` and fill in your values:

```bash
cp .env.example .env
```

| Variable | Description | Required |
|----------|-------------|----------|
| `DYNAMICS_BASE_URL` | Your Dynamics 365 instance URL | No (defaults to `https://microsoftsales.crm.dynamics.com`) |
| `DYNAMICS_CLIENT_ID` | Azure AD app registration client ID | **Yes** |
| `DYNAMICS_TENANT_ID` | Azure AD tenant ID | **Yes** |
| `DYNAMICS_USER_ID` | Your `systemuserid` GUID from Dynamics 365 | **Yes** |

### 3. Find your User ID

To find your `systemuserid`, run this OData query in your browser (while authenticated):

```
https://<your-instance>.crm.dynamics.com/api/data/v9.2/systemusers?$filter=internalemailaddress eq 'your.email@company.com'&$select=systemuserid,fullname
```

### 4. Configure GitHub Copilot CLI

Add to your `~/.copilot/mcp-config.json`:

```json
{
  "mcpServers": {
    "dynamics-crm": {
      "command": "python",
      "args": ["<full-path-to>/dynamics-mcp/server.py"],
      "env": {
        "DYNAMICS_CLIENT_ID": "your-client-id",
        "DYNAMICS_TENANT_ID": "your-tenant-id",
        "DYNAMICS_USER_ID": "your-user-id"
      },
      "tools": ["*"]
    }
  }
}
```

> **Tip:** You can also set the env vars in your shell profile or `.env` file instead of inline in the config.

### 5. First run

On first launch, the server will open a browser for interactive Azure AD login. After authenticating, the token is cached locally in `.token_cache.json`.

## Authentication

Uses MSAL (Microsoft Authentication Library) with the **public client** flow:
- Tokens are cached in `.token_cache.json` (auto-refreshed)
- First login requires interactive browser auth
- Subsequent runs use cached/refresh tokens silently

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

## License

Internal use only.

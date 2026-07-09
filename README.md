# graphkit

A small, dependency-light Python toolkit for authenticating to the
Microsoft Graph API using the **client credentials flow** (app-only
auth -- no user sign-in required), plus a thin HTTP client for making
calls against it.

This is useful for background scripts, scheduled jobs, or shared-mailbox
automations that need to act on Outlook mail, Planner tasks, SharePoint,
or any other Graph resource without a human logging in each time.

If this saves you time, consider [buying me a coffee](https://ko-fi.com/duckboard).

## Features

- OAuth2 client credentials flow via [MSAL](https://github.com/AzureAD/microsoft-authentication-library-for-python)
- No secrets in code -- credentials are read from environment variables
- A minimal `GraphClient` for GET/POST/PATCH/DELETE against any Graph endpoint
- Two worked examples: drafting an email, creating a Planner task

## Installation

```bash
pip install -r requirements.txt
```

or, to install as an editable package:

```bash
pip install -e .
```

## Azure app registration

You need an Azure AD (Entra ID) app registration before this will work.
If you don't already have one:

1. Go to [portal.azure.com](https://portal.azure.com) -> **Azure Active
   Directory** (or **Entra ID**) -> **App registrations** -> **New
   registration**.
2. Give it a name, leave the default options, and click **Register**.
3. Note down the **Application (client) ID** and **Directory (tenant)
   ID** from the app's Overview page.
4. Go to **Certificates & secrets** -> **New client secret**. Copy the
   secret **value** immediately -- it's only shown once.
5. Go to **API permissions** -> **Add a permission** -> **Microsoft
   Graph** -> **Application permissions**, and add whichever
   permissions your use case needs, for example:
   - `Mail.ReadWrite` (draft/read/send mail as any mailbox)
   - `Tasks.ReadWrite` (Planner tasks)
   - `Group.Read.All` (look up Microsoft 365 Groups by name)
   - `User.Read.All` (look up users by name/UPN)
6. Click **Grant admin consent** for your organisation (this step
   requires a Global Administrator or Application Administrator role).

Application permissions act on the whole tenant, scoped only by
whichever specific mailbox/site/group you address in your API calls --
there's no per-user consent screen for this flow. Only grant the
permissions you actually need.

## Configuration

Copy `.env.example` to `.env` and fill in your values:

```bash
cp .env.example .env
```

```
GRAPH_TENANT_ID=your-tenant-id
GRAPH_CLIENT_ID=your-client-id
GRAPH_CLIENT_SECRET=your-client-secret
```

`.env` is already in `.gitignore` -- never commit real credentials.

To load `.env` automatically in your own scripts, install
`python-dotenv` and add this before importing graphkit:

```python
from dotenv import load_dotenv
load_dotenv()
```

Alternatively, just set the three variables directly in your shell or
process environment -- graphkit doesn't require python-dotenv, it only
reads `os.environ`.

## Usage

```python
from graphkit import get_token, GraphClient

token = get_token()          # reads GRAPH_TENANT_ID / GRAPH_CLIENT_ID / GRAPH_CLIENT_SECRET
graph = GraphClient(token)

me = graph.get("/users/someone@yourdomain.com")
print(me["displayName"])
```

`GraphClient` is intentionally minimal -- `.get()`, `.post()`,
`.patch()`, `.delete()` -- so it works with any Graph endpoint, not
just the ones this package happens to have examples for.

### Examples

- `examples/send_email_draft.py` -- create a draft email in a mailbox's
  Drafts folder
- `examples/create_planner_task.py` -- look up a Microsoft 365 Group,
  Planner plan, and bucket by name, then create a task in it

Run either directly once you've set up your `.env`:

```bash
python examples/send_email_draft.py
python examples/create_planner_task.py
```

## Testing

```bash
pip install pytest
pytest
```

Tests mock MSAL, so they run offline without real Azure credentials.

## Why client credentials and not device code / interactive login?

Client credentials (app-only) auth is the right choice when a script
runs unattended -- no browser, no one available to click "allow" on a
sign-in prompt. If you need a script to act *as a specific signed-in
user* instead of a shared app identity, you'd want MSAL's device code
or interactive flows instead, which this package doesn't cover.

## License

MIT -- see [LICENSE](LICENSE). Do whatever you like with this; a
credit or a [Ko-fi tip](https://ko-fi.com/duckboard) is
appreciated but never required.

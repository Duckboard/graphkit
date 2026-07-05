"""
create_planner_task.py -- create a Microsoft Planner task via Graph API

Run with:  python create_planner_task.py

Demonstrates chaining several Graph lookups (group -> plan -> bucket ->
user) before creating a task -- a common pattern when your script
needs to find things by name rather than by ID. Copy this file and
adapt it for your own plan/bucket layout.

Requires an Azure app registration with Group.Read.All, Tasks.ReadWrite,
and User.Read.All Application permissions (admin consent required) --
see README.md.

Before running, set these environment variables:
    GRAPH_TENANT_ID
    GRAPH_CLIENT_ID
    GRAPH_CLIENT_SECRET
    PLANNER_GROUP_NAME   -- e.g. "Operations"
    PLANNER_PLAN_NAME    -- e.g. "Tasks"
    PLANNER_BUCKET_NAME  -- e.g. "To Do"
"""

import os
import sys
from datetime import date, timedelta

# Uncomment these two lines if you're using a .env file:
# from dotenv import load_dotenv
# load_dotenv()

from graphkit import get_token, GraphClient


def find_group_by_name(graph: GraphClient, name: str) -> dict:
    result = graph.get(f"/groups?$filter=displayName eq '{name}'&$select=id,displayName")
    match = next((g for g in result["value"] if g["displayName"] == name), None)
    if match:
        return match

    # Fall back to a case-insensitive partial match.
    result = graph.get("/groups?$select=id,displayName")
    for group in result["value"]:
        if name.lower() in group["displayName"].lower():
            return group

    raise RuntimeError(f"Could not find a group matching '{name}'")


def find_plan_by_name(graph: GraphClient, group_id: str, name: str) -> dict:
    result = graph.get(f"/groups/{group_id}/planner/plans")
    match = next((p for p in result["value"] if p.get("title", "").lower() == name.lower()), None)
    if not match:
        available = [p.get("title") for p in result["value"]]
        raise RuntimeError(f"Could not find plan '{name}'. Plans available: {available}")
    return match


def find_bucket_by_name(graph: GraphClient, plan_id: str, name: str) -> dict:
    result = graph.get(f"/planner/plans/{plan_id}/buckets")
    match = next((b for b in result["value"] if b["name"].lower() == name.lower()), None)
    if not match:
        available = [b["name"] for b in result["value"]]
        raise RuntimeError(f"Could not find bucket '{name}'. Buckets available: {available}")
    return match


def next_weekday(weekday: int) -> str:
    """Return an ISO datetime string for the next occurrence of `weekday`
    (0 = Monday ... 6 = Sunday), always at least one day in the future."""
    today = date.today()
    days_ahead = (weekday - today.weekday()) % 7
    if days_ahead == 0:
        days_ahead = 7
    return (today + timedelta(days=days_ahead)).strftime("%Y-%m-%dT00:00:00Z")


def main():
    group_name = os.environ.get("PLANNER_GROUP_NAME")
    plan_name = os.environ.get("PLANNER_PLAN_NAME")
    bucket_name = os.environ.get("PLANNER_BUCKET_NAME")
    if not (group_name and plan_name and bucket_name):
        print("Set PLANNER_GROUP_NAME, PLANNER_PLAN_NAME, and PLANNER_BUCKET_NAME.")
        sys.exit(1)

    token = get_token()
    graph = GraphClient(token)

    group = find_group_by_name(graph, group_name)
    print(f"Group:  {group['displayName']} ({group['id']})")

    plan = find_plan_by_name(graph, group["id"], plan_name)
    print(f"Plan:   {plan['title']} ({plan['id']})")

    bucket = find_bucket_by_name(graph, plan["id"], bucket_name)
    print(f"Bucket: {bucket['name']} ({bucket['id']})")

    # --- Replace this with your own task details ---
    task_body = {
        "planId": plan["id"],
        "bucketId": bucket["id"],
        "title": "Example task -- edit me",
        "dueDateTime": next_weekday(4),  # next Friday
    }

    response = graph.post("/planner/tasks", task_body)
    if response.status_code == 201:
        print(f"\nTask created -> id {response.json()['id']}")
    else:
        print(f"\nFailed to create task ({response.status_code}): {GraphClient.error_message(response)}")


if __name__ == "__main__":
    main()

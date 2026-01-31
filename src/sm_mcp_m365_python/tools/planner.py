"""Planner tools for M365 MCP Server."""

from typing import Any, Dict, List, Optional

from mcp.types import Tool

from ..auth import M365OAuth
from ..graph import GraphClient


PLANNER_TOOLS: List[Tool] = [
    Tool(
        name="m365_list_teams",
        description="List Microsoft Teams the user has joined",
        inputSchema={
            "type": "object",
            "properties": {},
            "required": [],
        },
    ),
    Tool(
        name="m365_list_plans",
        description="List all Planner plans accessible to the user, optionally filtered by team",
        inputSchema={
            "type": "object",
            "properties": {
                "team_id": {
                    "type": "string",
                    "description": "Optional team/group ID to filter plans",
                },
            },
            "required": [],
        },
    ),
    Tool(
        name="m365_get_plan",
        description="Get details of a specific Planner plan including buckets",
        inputSchema={
            "type": "object",
            "properties": {
                "plan_id": {
                    "type": "string",
                    "description": "The plan ID",
                },
            },
            "required": ["plan_id"],
        },
    ),
    Tool(
        name="m365_list_tasks",
        description="List tasks in a Planner plan or all tasks assigned to me",
        inputSchema={
            "type": "object",
            "properties": {
                "plan_id": {
                    "type": "string",
                    "description": "Plan ID to list tasks from. If not provided, lists all my tasks.",
                },
            },
            "required": [],
        },
    ),
    Tool(
        name="m365_get_task",
        description="Get details of a specific Planner task including description and checklist",
        inputSchema={
            "type": "object",
            "properties": {
                "task_id": {
                    "type": "string",
                    "description": "The task ID",
                },
            },
            "required": ["task_id"],
        },
    ),
    Tool(
        name="m365_create_task",
        description="Create a new Planner task",
        inputSchema={
            "type": "object",
            "properties": {
                "plan_id": {
                    "type": "string",
                    "description": "The plan ID to create the task in",
                },
                "title": {
                    "type": "string",
                    "description": "Task title",
                },
                "bucket_id": {
                    "type": "string",
                    "description": "Optional bucket ID",
                },
                "due_date": {
                    "type": "string",
                    "description": "Due date in ISO 8601 format (e.g., 2026-02-15T00:00:00Z)",
                },
                "priority": {
                    "type": "integer",
                    "description": "Priority: 1=urgent, 3=important, 5=medium, 9=low",
                    "enum": [1, 3, 5, 9],
                },
            },
            "required": ["plan_id", "title"],
        },
    ),
    Tool(
        name="m365_update_task",
        description="Update a Planner task (mark complete, change title, etc.)",
        inputSchema={
            "type": "object",
            "properties": {
                "task_id": {
                    "type": "string",
                    "description": "The task ID",
                },
                "title": {
                    "type": "string",
                    "description": "New task title",
                },
                "percent_complete": {
                    "type": "integer",
                    "description": "Completion percentage: 0 (not started), 50 (in progress), 100 (complete)",
                    "enum": [0, 50, 100],
                },
                "due_date": {
                    "type": "string",
                    "description": "Due date in ISO 8601 format",
                },
                "priority": {
                    "type": "integer",
                    "description": "Priority: 1=urgent, 3=important, 5=medium, 9=low",
                    "enum": [1, 3, 5, 9],
                },
                "bucket_id": {
                    "type": "string",
                    "description": "Move to different bucket",
                },
            },
            "required": ["task_id"],
        },
    ),
    Tool(
        name="m365_delete_task",
        description="Delete a Planner task",
        inputSchema={
            "type": "object",
            "properties": {
                "task_id": {
                    "type": "string",
                    "description": "The task ID to delete",
                },
            },
            "required": ["task_id"],
        },
    ),
]


async def handle_list_teams(
    arguments: Dict[str, Any],
    oauth: M365OAuth,
    client: GraphClient,
) -> Dict[str, Any]:
    """Handle m365_list_teams tool call."""
    teams = await client.list_joined_teams()

    return {
        "teams": [
            {
                "id": team.get("id"),
                "displayName": team.get("displayName"),
                "description": team.get("description"),
            }
            for team in teams
        ],
        "count": len(teams),
    }


async def handle_list_plans(
    arguments: Dict[str, Any],
    oauth: M365OAuth,
    client: GraphClient,
) -> Dict[str, Any]:
    """Handle m365_list_plans tool call."""
    team_id = arguments.get("team_id")

    if team_id:
        plans = await client.list_group_plans(team_id)
    else:
        plans = await client.list_my_plans()

    return {
        "plans": [
            {
                "id": plan.get("id"),
                "title": plan.get("title"),
                "owner": plan.get("owner"),
                "createdDateTime": plan.get("createdDateTime"),
            }
            for plan in plans
        ],
        "count": len(plans),
    }


async def handle_get_plan(
    arguments: Dict[str, Any],
    oauth: M365OAuth,
    client: GraphClient,
) -> Dict[str, Any]:
    """Handle m365_get_plan tool call."""
    plan_id = arguments["plan_id"]

    plan = await client.get_plan(plan_id)
    buckets = await client.list_buckets(plan_id)

    return {
        "plan": {
            "id": plan.get("id"),
            "title": plan.get("title"),
            "owner": plan.get("owner"),
            "createdDateTime": plan.get("createdDateTime"),
        },
        "buckets": [
            {
                "id": bucket.get("id"),
                "name": bucket.get("name"),
                "orderHint": bucket.get("orderHint"),
            }
            for bucket in buckets
        ],
    }


async def handle_list_tasks(
    arguments: Dict[str, Any],
    oauth: M365OAuth,
    client: GraphClient,
) -> Dict[str, Any]:
    """Handle m365_list_tasks tool call."""
    plan_id = arguments.get("plan_id")

    if plan_id:
        tasks = await client.list_plan_tasks(plan_id)
    else:
        tasks = await client.list_my_tasks()

    # Map priority numbers to labels
    priority_map = {1: "urgent", 3: "important", 5: "medium", 9: "low"}

    return {
        "tasks": [
            {
                "id": task.get("id"),
                "title": task.get("title"),
                "bucketId": task.get("bucketId"),
                "percentComplete": task.get("percentComplete"),
                "priority": priority_map.get(task.get("priority", 5), "medium"),
                "dueDateTime": task.get("dueDateTime"),
                "createdDateTime": task.get("createdDateTime"),
                "assigneePriority": task.get("assigneePriority"),
                "@odata.etag": task.get("@odata.etag"),
            }
            for task in tasks
        ],
        "count": len(tasks),
    }


async def handle_get_task(
    arguments: Dict[str, Any],
    oauth: M365OAuth,
    client: GraphClient,
) -> Dict[str, Any]:
    """Handle m365_get_task tool call."""
    task_id = arguments["task_id"]

    task = await client.get_task(task_id)
    details = await client.get_task_details(task_id)

    priority_map = {1: "urgent", 3: "important", 5: "medium", 9: "low"}

    # Extract checklist items
    checklist = details.get("checklist", {})
    checklist_items = [
        {
            "id": item_id,
            "title": item.get("title"),
            "isChecked": item.get("isChecked"),
        }
        for item_id, item in checklist.items()
    ]

    return {
        "task": {
            "id": task.get("id"),
            "title": task.get("title"),
            "bucketId": task.get("bucketId"),
            "planId": task.get("planId"),
            "percentComplete": task.get("percentComplete"),
            "priority": priority_map.get(task.get("priority", 5), "medium"),
            "dueDateTime": task.get("dueDateTime"),
            "startDateTime": task.get("startDateTime"),
            "createdDateTime": task.get("createdDateTime"),
            "completedDateTime": task.get("completedDateTime"),
            "assignments": list(task.get("assignments", {}).keys()),
            "@odata.etag": task.get("@odata.etag"),
        },
        "details": {
            "description": details.get("description", ""),
            "checklist": checklist_items,
            "references": list(details.get("references", {}).keys()),
        },
    }


async def handle_create_task(
    arguments: Dict[str, Any],
    oauth: M365OAuth,
    client: GraphClient,
) -> Dict[str, Any]:
    """Handle m365_create_task tool call."""
    plan_id = arguments["plan_id"]
    title = arguments["title"]
    bucket_id = arguments.get("bucket_id")
    due_date = arguments.get("due_date")
    priority = arguments.get("priority")

    task = await client.create_task(
        plan_id=plan_id,
        title=title,
        bucket_id=bucket_id,
        due_date=due_date,
        priority=priority,
    )

    return {
        "created": True,
        "task": {
            "id": task.get("id"),
            "title": task.get("title"),
            "bucketId": task.get("bucketId"),
            "planId": task.get("planId"),
        },
    }


async def handle_update_task(
    arguments: Dict[str, Any],
    oauth: M365OAuth,
    client: GraphClient,
) -> Dict[str, Any]:
    """Handle m365_update_task tool call."""
    task_id = arguments["task_id"]

    # First get the task to get the etag
    task = await client.get_task(task_id)
    etag = task.get("@odata.etag")

    if not etag:
        return {"error": "Could not get task etag for update"}

    result = await client.update_task(
        task_id=task_id,
        etag=etag,
        title=arguments.get("title"),
        bucket_id=arguments.get("bucket_id"),
        percent_complete=arguments.get("percent_complete"),
        due_date=arguments.get("due_date"),
        priority=arguments.get("priority"),
    )

    return {
        "updated": True,
        "task_id": task_id,
    }


async def handle_delete_task(
    arguments: Dict[str, Any],
    oauth: M365OAuth,
    client: GraphClient,
) -> Dict[str, Any]:
    """Handle m365_delete_task tool call."""
    task_id = arguments["task_id"]

    # First get the task to get the etag
    task = await client.get_task(task_id)
    etag = task.get("@odata.etag")

    if not etag:
        return {"error": "Could not get task etag for delete"}

    await client.delete_task(task_id, etag)

    return {
        "deleted": True,
        "task_id": task_id,
    }


PLANNER_HANDLERS = {
    "m365_list_teams": handle_list_teams,
    "m365_list_plans": handle_list_plans,
    "m365_get_plan": handle_get_plan,
    "m365_list_tasks": handle_list_tasks,
    "m365_get_task": handle_get_task,
    "m365_create_task": handle_create_task,
    "m365_update_task": handle_update_task,
    "m365_delete_task": handle_delete_task,
}

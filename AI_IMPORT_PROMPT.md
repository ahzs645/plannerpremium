# AI Import Prompt Guide

Use these prompts with any AI (ChatGPT, Claude, Copilot, etc.) to generate task lists you can import into Planner Premium.

---

## Microsoft To Do — Import Prompt

Copy and paste the following into any AI, replacing the last line with your project description:

```
Generate a task list in CSV format with these exact columns:

Title,Priority,DueDate,Description,ChecklistItems

Rules:
- Title (required): The task name
- Priority: "High", "Normal", or "Low"
- DueDate: YYYY-MM-DD format
- Description: Plain text notes for the task
- ChecklistItems: Subtasks separated by semicolons (;)
- Wrap any field containing commas in double quotes
- Do NOT include any markdown formatting, code fences, or extra text — output ONLY the raw CSV

Example output:
Title,Priority,DueDate,Description,ChecklistItems
Buy groceries,High,2025-01-25,Weekly shopping,Milk;Bread;Eggs;Butter
Call dentist,Normal,2025-01-26,Schedule appointment,
Finish report,High,2025-01-24,Q4 quarterly report,Review data;Write summary;Add charts

Now generate a task list for: [DESCRIBE YOUR PROJECT OR GOAL HERE]
```

### Field Reference

| Column         | Required | Values                          | Notes                              |
| -------------- | -------- | ------------------------------- | ---------------------------------- |
| Title          | Yes      | Any text                        | The task name                      |
| Priority       | No       | `High`, `Normal`, `Low`         | Defaults to `Normal` if omitted    |
| DueDate        | No       | `YYYY-MM-DD`                    | e.g. `2025-03-15`                  |
| Description    | No       | Any text                        | Plain text notes                   |
| ChecklistItems | No       | Semicolon-separated list        | e.g. `Step 1;Step 2;Step 3`       |

---

## Microsoft Planner — Import Prompt

Copy and paste the following into any AI, replacing the last line with your project description:

```
Generate a task list in CSV format with these exact columns:

OutlineNumber,Title,Bucket,Priority,StartDate,DueDate,AssignedTo,Description,ChecklistItems

Rules:
- OutlineNumber (required): Dotted hierarchy numbers (1, 1.1, 1.1.1, 2, 2.1, etc.)
  - Top-level tasks: 1, 2, 3
  - Subtasks: 1.1, 1.2, 2.1
  - Sub-subtasks: 1.1.1, 1.1.2
- Title (required): The task name
- Bucket: Group/category name for the task
- Priority: "Urgent", "High", "Medium", or "Low"
- StartDate/DueDate: YYYY-MM-DD format
- AssignedTo: Email addresses separated by semicolons
- Description: Plain text notes
- ChecklistItems: Subtasks separated by semicolons (;)
- Wrap any field containing commas in double quotes
- Do NOT include any markdown formatting, code fences, or extra text — output ONLY the raw CSV

Example output:
OutlineNumber,Title,Bucket,Priority,StartDate,DueDate,AssignedTo,Description,ChecklistItems
1,Phase 1: Planning,Backlog,High,2025-01-20,2025-01-31,,Project planning phase,
1.1,Define requirements,Backlog,High,2025-01-20,2025-01-22,pm@company.com,Gather requirements,Interview stakeholders;Document requirements
1.2,Create timeline,Backlog,Medium,2025-01-23,2025-01-25,pm@company.com,,
2,Phase 2: Development,Sprint 1,High,2025-02-01,2025-02-28,,Development phase,
2.1,Setup dev environment,Sprint 1,Urgent,2025-02-01,2025-02-03,dev@company.com,,Install tools;Configure CI/CD
2.2,Build feature A,Sprint 1,High,2025-02-04,2025-02-14,dev@company.com,Core feature,Design;Implement;Test

Now generate a task list for: [DESCRIBE YOUR PROJECT OR GOAL HERE]
```

### Field Reference

| Column        | Required | Values                              | Notes                                    |
| ------------- | -------- | ----------------------------------- | ---------------------------------------- |
| OutlineNumber | Yes      | Dotted numbers (`1`, `1.1`, `1.1.1`)| Defines parent-child hierarchy           |
| Title         | Yes      | Any text                            | The task name                            |
| Bucket        | No       | Any text                            | Must match an existing bucket in Planner |
| Priority      | No       | `Urgent`, `High`, `Medium`, `Low`   | Defaults to `Medium` if omitted          |
| StartDate     | No       | `YYYY-MM-DD`                        | e.g. `2025-03-01`                        |
| DueDate       | No       | `YYYY-MM-DD`                        | e.g. `2025-03-15`                        |
| AssignedTo    | No       | Semicolon-separated emails          | e.g. `a@co.com;b@co.com`                |
| Description   | No       | Any text                            | Plain text notes                         |
| ChecklistItems| No       | Semicolon-separated list            | e.g. `Step 1;Step 2;Step 3`             |

---

## How to Import

1. Copy the AI-generated CSV output
2. Paste it into a text editor and save as a `.csv` file
3. Open Planner Premium and go to the **Import** page
4. Select your CSV file and follow the on-screen steps

## Tips

- Tell the AI how many tasks you want (e.g. "generate 20 tasks")
- Be specific about your project so the AI gives relevant task names and descriptions
- If the AI wraps the output in code fences (` ``` `), remove them before saving
- Review and adjust dates/priorities before importing

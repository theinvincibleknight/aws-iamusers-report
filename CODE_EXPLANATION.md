# Code Explanation

Detailed walkthrough of `lambda_function.py` — the AWS Lambda function that generates a multi-account IAM Users Inventory Report.

## Table of Contents

- [Overview](#overview)
- [Imports and Configuration](#imports-and-configuration)
- [Core Functions](#core-functions)
  - [assume_role()](#assume_role)
  - [get_user_policies()](#get_user_policies)
  - [get_access_key_details()](#get_access_key_details)
  - [get_last_activity()](#get_last_activity)
  - [get_console_access_status()](#get_console_access_status)
  - [get_user_creation_time()](#get_user_creation_time)
  - [collect_iam_data()](#collect_iam_data)
  - [create_excel_report()](#create_excel_report)
  - [upload_to_s3()](#upload_to_s3)
  - [lambda_handler()](#lambda_handler)
- [Data Flow](#data-flow)
- [Error Handling](#error-handling)

---

## Overview

The function iterates over a dictionary of AWS accounts, assumes a read-only role in each one, and collects comprehensive IAM user data including:

- Attached policies (AWS managed, customer managed, inline)
- Policies inherited from IAM groups
- Access key details (ID, status, age, last used)
- Console access status
- User activity and creation timestamps

All data is compiled into a styled Excel workbook and uploaded to S3.

---

## Imports and Configuration

```python
import boto3
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime, timezone
from io import BytesIO
import logging
```

| Import | Purpose |
|--------|---------|
| `boto3` | AWS SDK — used for IAM, STS, and S3 API calls |
| `openpyxl` | Excel file generation (provided via Lambda Layer) |
| `openpyxl.styles` | Cell formatting — fonts, colors, borders, alignment |
| `datetime` | Timestamp handling and key age calculation |
| `BytesIO` | In-memory file buffer for S3 upload (no disk write needed) |
| `logging` | CloudWatch Logs integration |

### Configuration Constants

```python
S3_BUCKET = 'aws-artifact-collector'

ACCOUNTS = {
    'dev': {'role_arn': 'arn:aws:iam::058264257786:role/inventory-readonly-role'},
    'uat': {'role_arn': 'arn:aws:iam::891377165721:role/inventory-readonly-role'},
    ...
}
```

- `S3_BUCKET`: Target bucket for the report
- `ACCOUNTS`: Dictionary mapping friendly account names to their cross-account role ARNs

---

## Core Functions

### assume_role()

```python
def assume_role(role_arn, session_name='IAMInventorySession'):
```

**Purpose:** Assumes a cross-account IAM role and returns a new `boto3.Session` with temporary credentials.

**How it works:**
1. Calls `sts:AssumeRole` with the target role ARN
2. Extracts temporary credentials (AccessKeyId, SecretAccessKey, SessionToken)
3. Creates and returns a new `boto3.Session` using those credentials
4. Session is valid for 1 hour (`DurationSeconds=3600`)

**Why a session?** Creating a session (instead of a client) allows us to create multiple service clients (IAM, etc.) from the same assumed-role credentials.

---

### get_user_policies()

```python
def get_user_policies(iam_client, username):
```

**Purpose:** Retrieves all policies effective for a user — from three sources.

**Policy sources collected:**

| Source | API Call | Category |
|--------|----------|----------|
| Directly attached managed policies | `list_attached_user_policies` | AWS Managed or Customer Managed |
| Directly attached inline policies | `list_user_policies` | Inline |
| Group-attached managed policies | `list_attached_group_policies` | AWS Managed or Customer Managed |
| Group inline policies | `list_group_policies` | Inline |

**Classification logic:**
- If the policy ARN starts with `arn:aws:iam::aws:policy` → AWS Managed
- Otherwise → Customer Managed

**Group inheritance:** The function first calls `list_groups_for_user` to get all groups, then iterates each group to collect its policies. Group-inherited policies are labeled with `(via group: GroupName)` for clarity.

**Pagination:** All list operations use paginators to handle accounts with many policies.

---

### get_access_key_details()

```python
def get_access_key_details(iam_client, username):
```

**Purpose:** Gets all access keys for a user with their metadata.

**Data collected per key:**

| Field | Source | Notes |
|-------|--------|-------|
| Key ID | `list_access_keys` | The AKIA... identifier |
| Status | `list_access_keys` | Active or Inactive |
| Creation Date | `list_access_keys` | When the key was created |
| Key Age (Days) | Calculated | `(now - create_date).days`, only for Active keys |
| Last Used Date | `get_access_key_last_used` | Last API call timestamp |
| Last Used Service | `get_access_key_last_used` | e.g., s3, ec2, iam |

**Key age calculation:** Only computed for Active keys. Inactive keys show `N/A (Inactive)` since age is not a security concern for disabled keys.

---

### get_last_activity()

```python
def get_last_activity(iam_client, username):
```

**Purpose:** Returns the timestamp of the user's last console login.

**How it works:**
- Calls `iam:GetUser` and checks for `PasswordLastUsed` field
- If present, returns the formatted timestamp
- If absent (user never logged in via console), returns `'No console activity'`

---

### get_console_access_status()

```python
def get_console_access_status(iam_client, username):
```

**Purpose:** Determines if a user can log into the AWS Console.

**How it works:**
- Calls `iam:GetLoginProfile` for the user
- If successful → user has a password → returns `'Enabled'`
- If `NoSuchEntityException` → no login profile → returns `'Disabled'`

This is the definitive check — a user without a login profile cannot access the console regardless of their policies.

---

### get_user_creation_time()

```python
def get_user_creation_time(iam_client, username):
```

**Purpose:** Returns when the IAM user was created.

**How it works:** Calls `iam:GetUser` and formats the `CreateDate` field.

---

### collect_iam_data()

```python
def collect_iam_data(account_name, session):
```

**Purpose:** Orchestrates data collection for a single account. This is the main per-account workhorse.

**Flow:**
1. Creates an IAM client from the assumed-role session
2. Lists all users using a paginator
3. For each user, calls all the helper functions above
4. Builds one row per access key (or one row if user has no keys)

**Multiple access keys handling:**
- AWS allows up to 2 access keys per user
- Users with multiple keys get multiple rows in the report
- Policy columns are repeated on each row for readability in Excel

---

### create_excel_report()

```python
def create_excel_report(all_data):
```

**Purpose:** Transforms the collected data list into a styled Excel workbook.

**Styling applied:**

| Element | Style |
|---------|-------|
| Header row | Blue background (#4472C4), white bold font, centered |
| All cells | Thin borders, wrap text enabled, top-aligned |
| Inactive keys | Red font color |
| Keys > 90 days old | Orange bold font |
| Disabled console | Gray font color |

**Features:**
- `Sr. No` column auto-increments starting from 1
- `wrap_text=True` on all cells keeps rows compact
- Column widths auto-adjust (sampled from first 50 rows, capped at 35 chars)
- Header row is frozen (`freeze_panes = 'A2'`) for scrolling

---

### upload_to_s3()

```python
def upload_to_s3(workbook):
```

**Purpose:** Saves the workbook to S3 with a timestamped path.

**Path format:**
```
iam_users/<year>/<month>/iam_users_report_<year>_<month>_<HHMMSS>.xlsx
```

**How it works:**
1. Saves the workbook to a `BytesIO` buffer (in-memory, no `/tmp` disk usage)
2. Calls `s3:PutObject` with the correct content type
3. Returns the full S3 URI for logging

---

### lambda_handler()

```python
def lambda_handler(event, context):
```

**Purpose:** Entry point for the Lambda function. Orchestrates the entire flow.

**Flow:**
1. Iterates over all accounts in `ACCOUNTS`
2. Assumes role in each account
3. Collects IAM data (skips account if role assumption fails)
4. Aggregates all records
5. Generates Excel report
6. Uploads to S3
7. Returns success response with the S3 path

---

## Data Flow

```
lambda_handler()
  │
  ├─▶ For each account in ACCOUNTS:
  │     │
  │     ├─▶ assume_role(role_arn)
  │     │     └─▶ Returns boto3.Session with temp credentials
  │     │
  │     └─▶ collect_iam_data(account_name, session)
  │           │
  │           ├─▶ list_users (paginated)
  │           │
  │           └─▶ For each user:
  │                 ├─▶ get_user_policies()
  │                 │     ├─▶ list_attached_user_policies
  │                 │     ├─▶ list_user_policies (inline)
  │                 │     └─▶ list_groups_for_user
  │                 │           ├─▶ list_attached_group_policies
  │                 │           └─▶ list_group_policies
  │                 │
  │                 ├─▶ get_access_key_details()
  │                 │     ├─▶ list_access_keys
  │                 │     └─▶ get_access_key_last_used (per key)
  │                 │
  │                 ├─▶ get_last_activity()
  │                 ├─▶ get_console_access_status()
  │                 └─▶ get_user_creation_time()
  │
  ├─▶ create_excel_report(all_data)
  │     └─▶ Returns openpyxl Workbook
  │
  └─▶ upload_to_s3(workbook)
        └─▶ Returns S3 URI
```

---

## Error Handling

The function uses a **graceful degradation** approach:

| Scenario | Behavior |
|----------|----------|
| Cannot assume role for an account | Logs error, skips account, continues with others |
| Cannot get policies for a user | Logs error, returns empty lists for that user |
| Cannot get access key last used | Defaults to `'Never'` |
| Cannot check console access | Returns `'Error'` in that cell |
| No data collected from any account | Returns success with informational message |

All errors are logged to CloudWatch via the `logger` for troubleshooting without breaking the overall report generation.

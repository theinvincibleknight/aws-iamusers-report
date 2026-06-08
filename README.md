# AWS Multi-Account IAM Users Inventory Report Lambda

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

Lambda function that collects IAM user details from multiple AWS accounts using cross-account IAM role assumption, generates a styled Excel report, and stores it in S3.

No access keys. No secrets to rotate. Uses `sts:AssumeRole` for secure cross-account access with temporary credentials.

## Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                    Central Account (Main)                         │
│                                                                   │
│  ┌──────────────┐     ┌──────────────┐     ┌─────────────────┐  │
│  │  EventBridge │────▶│    Lambda     │────▶│   S3 Bucket     │  │
│  │  (Monthly)   │     │  Function    │     │  (Excel Report) │  │
│  └──────────────┘     └──────┬───────┘     └─────────────────┘  │
│                              │                                    │
└──────────────────────────────┼────────────────────────────────────┘
                               │ sts:AssumeRole
          ┌────────────────────┼────────────────────┐
          ▼                    ▼                    ▼
   ┌─────────────┐     ┌─────────────┐     ┌─────────────┐
   │  Dev Account │     │  UAT Account │     │ Prod Account │
   │  (ReadOnly)  │     │  (ReadOnly)  │     │  (ReadOnly)  │
   └─────────────┘     └─────────────┘     └─────────────┘
```

### Execution Flow

```
1. EventBridge triggers Lambda (scheduled or manual)
2. Lambda calls sts:AssumeRole for each target account
3. Gets temporary credentials (no stored secrets)
4. Lists all IAM users in the account
5. Collects per-user:
   - AWS Managed Policies (direct + via groups)
   - Customer Managed Policies (direct + via groups)
   - Inline Policies (direct + via groups)
   - Last Activity (password last used)
   - Access Key IDs, Status, Age, Last Used, Creation Time
   - Console Access status (login profile check)
   - User Creation Time
6. Generates styled Excel report with all data
7. Uploads to s3://aws-artifact-collector/iam_users/YYYY/MM/iam_users_report_YYYY_MM_HHMMSS.xlsx
8. Returns summary with record counts
```

## Data Collected

| Column | Description |
|--------|-------------|
| Sr. No | Serial number |
| Account | Account alias (dev, uat, prod, etc.) |
| UserName | IAM username |
| AWS Managed Policies | AWS-managed policies attached directly or via groups |
| Customer Managed Policies | Customer-managed policies attached directly or via groups |
| Inline Policies | Inline policies on user or inherited from groups |
| Last Activity | Timestamp of last console login (password last used) |
| Access Key ID | Access key identifier (one row per key) |
| Access Key Status | Active / Inactive |
| Active Key Age (Days) | Days since key creation (active keys only) |
| Access Key Last Used | Timestamp of last API call with this key |
| Access Key Last Used Service | AWS service last accessed with this key |
| Access Key Creation Time | When the access key was created |
| Console Access | Enabled / Disabled (login profile exists or not) |
| User Creation Time | When the IAM user was created |

> **Note:** Users with multiple access keys get one row per key. Policies are repeated for readability.

## Setup

### Step 1: Create IAM Role in Each Target Account

Run this in each target account (dev, uat, prod, oldprod, network, sharedservice). Replace `XXXXXXXXXXXX` with your central/main account ID.

Create the trust policy file:

```bash
cat > trust-policy.json << 'EOF'
{
  "Version": "2012-10-17",
  "Statement": [
    {
      "Effect": "Allow",
      "Principal": {
        "AWS": "arn:aws:iam::XXXXXXXXXXXX:root"
      },
      "Action": "sts:AssumeRole"
    }
  ]
}
EOF
```

Create the role and attach ReadOnlyAccess:

```bash
aws iam create-role \
  --role-name inventory-readonly-role \
  --assume-role-policy-document file://trust-policy.json

aws iam attach-role-policy \
  --role-name inventory-readonly-role \
  --policy-arn arn:aws:iam::aws:policy/ReadOnlyAccess
```

Repeat in each target account. The role name must be `inventory-readonly-role` in all accounts (or update the `ACCOUNTS` dict in the code).

### Step 2: Create S3 Bucket

1. Go to **S3 Console** → **Create bucket**
2. Bucket name: `aws-artifact-collector`
3. Region: Your preferred region
4. Keep default settings → **Create bucket**

### Step 3: Create Lambda Layer (if you don't have one already)

> **Note:** If you already have an `openpyxl` Lambda Layer from a previous function, skip this step and reuse it.

Open **AWS CloudShell** and run:

```bash
mkdir -p python && \
pip install openpyxl --no-deps -t python/ --quiet && \
pip install et-xmlfile -t python/ --quiet && \
zip -r openpyxl-layer.zip python/ && \
aws lambda publish-layer-version \
  --layer-name openpyxl \
  --zip-file fileb://openpyxl-layer.zip \
  --compatible-runtimes python3.11 python3.12 \
  --region ap-south-1
```

### Step 4: Create Lambda Function

1. Go to **Lambda Console** → **Create function** → **Author from scratch**
2. Configure:

   | Setting | Value |
   |---------|-------|
   | Function name | `iam-users-inventory-report` |
   | Runtime | Python 3.11 or 3.12 |
   | Architecture | x86_64 |

3. Click **Create function**
4. Under **Code** tab → paste the entire contents of `lambda_function.py` → click **Deploy**
5. Under **Configuration** → **General configuration** → **Edit**:
   - Timeout: **5 min** (increase if you have many users across accounts)
   - Memory: **256 MB**
6. Under **Code** tab → scroll to **Layers** → **Add a layer** → **Custom layers** → select `openpyxl` → **Add**

### Step 5: Lambda Execution Role (Inline Policy)

1. Go to **Configuration** → **Permissions** → click the **Role name**
2. **Add permissions** → **Create inline policy** → **JSON** tab:

```json
{
  "Version": "2012-10-17",
  "Statement": [
    {
      "Sid": "AssumeRoleInTargetAccounts",
      "Effect": "Allow",
      "Action": "sts:AssumeRole",
      "Resource": [
        "arn:aws:iam::111111111111:role/inventory-readonly-role",
        "arn:aws:iam::222222222222:role/inventory-readonly-role",
        "arn:aws:iam::333333333333:role/inventory-readonly-role",
        "arn:aws:iam::444444444444:role/inventory-readonly-role",
        "arn:aws:iam::555555555555:role/inventory-readonly-role",
        "arn:aws:iam::666666666666:role/inventory-readonly-role"
      ]
    },
    {
      "Sid": "S3UploadReports",
      "Effect": "Allow",
      "Action": "s3:PutObject",
      "Resource": "arn:aws:s3:::aws-artifact-collector/iam_users/*"
    }
  ]
}
```

3. Name it `iam-users-report-lambda-policy` → **Create policy**

## Invoke

### Manual (from Lambda Console)

Go to **Test** tab → Create test event with empty payload:

```json
{}
```

Click **Test**. This will process all accounts.

### Scheduled (EventBridge)

Go to **EventBridge** → **Schedules** → **Create schedule**:

| Schedule Name | Cron Expression | Target | Input |
|---------------|-----------------|--------|-------|
| iam-users-report-monthly | `cron(0 2 1 * ? *)` | iam-users-inventory-report | `{}` |

This runs monthly on the 1st at 2:00 AM UTC.

## S3 Output

```
s3://aws-artifact-collector/
  iam_users/
    2026/
      06/
        iam_users_report_2026_06_020015.xlsx
      07/
        iam_users_report_2026_07_020012.xlsx
```

## Excel Report Features

- **Styled headers** with blue background and white bold text
- **Wrap text** enabled on all cells for compact row height
- **Color coding:**
  - 🔴 Red text for Inactive access keys
  - 🟠 Orange bold text for Active keys older than 90 days
  - ⚪ Gray text for Disabled console access
- **Frozen header row** for easy scrolling
- **Auto-adjusted column widths** (capped at 35 chars, wrap handles overflow)
- **Group-inherited policies** labeled with `(via group: GroupName)`

## Adding New Accounts

1. In the new target account, create the role:
   ```bash
   aws iam create-role --role-name inventory-readonly-role \
     --assume-role-policy-document file://trust-policy.json
   aws iam attach-role-policy --role-name inventory-readonly-role \
     --policy-arn arn:aws:iam::aws:policy/ReadOnlyAccess
   ```

2. In `lambda_function.py`, add to `ACCOUNTS` dict:
   ```python
   'newaccount': {
       'role_arn': 'arn:aws:iam::777777777777:role/inventory-readonly-role',
   },
   ```

3. Update the Lambda inline policy to include the new role ARN in the `Resource` array.

4. Deploy the updated code.

## Project Files

```
├── lambda_function.py    # Lambda function code (paste into Lambda console)
├── README.md             # This setup guide
```

## Troubleshooting

| Error | Cause | Fix |
|-------|-------|-----|
| `No module named 'openpyxl'` | Layer not attached | Add the openpyxl layer to the function |
| `AccessDenied` on AssumeRole | Trust policy missing or wrong account ID | Check trust policy in target account allows your central account |
| `AccessDenied` on PutObject | Missing S3 permission in Lambda role | Add `s3:PutObject` for the bucket in inline policy |
| `Task timed out` | Too many users across accounts | Increase timeout to 10-15 min and memory to 512 MB |
| `is not authorized to perform: sts:AssumeRole` | Lambda role missing AssumeRole permission | Add the target role ARN to the inline policy Resource array |
| Empty report | No IAM users found | Verify users exist in target accounts; check CloudWatch logs |
| Console Access shows 'Error' | Permissions issue | Ensure `inventory-readonly-role` has `iam:GetLoginProfile` (included in ReadOnlyAccess) |

## IAM Permissions Required on Target Role

The `inventory-readonly-role` needs these IAM permissions (all included in `ReadOnlyAccess` managed policy):

- `iam:ListUsers`
- `iam:GetUser`
- `iam:ListAttachedUserPolicies`
- `iam:ListUserPolicies`
- `iam:ListGroupsForUser`
- `iam:ListAttachedGroupPolicies`
- `iam:ListGroupPolicies`
- `iam:ListAccessKeys`
- `iam:GetAccessKeyLastUsed`
- `iam:GetLoginProfile`

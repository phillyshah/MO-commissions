# Signing service — setup in Lenny's AWS account

This is the one piece of the approval chain the builder does not control (SIG-0). It runs in an AWS
account Lenny owns. Andy's IAM user can call it and write audit anchors; it cannot read the key, change
the function, or sign anything without a verified passkey assertion.

Region used below: `us-east-1`. Change consistently if another is preferred.

## 1. Account (Lenny, once)

1. Create the AWS account (or use an existing one Lenny administers). Turn on MFA for the root user
   and never use root again after step 2.
2. Create an IAM user `lenny-admin` with the `AdministratorAccess` policy and MFA. Do the rest signed
   in as that user.
3. Create an IAM user `maxx-ap-portal` for the portal (Andy's). **Programmatic access only, no console
   password.** Attach exactly the policy in §6 — nothing else, no groups.

## 2. KMS key

Console → KMS → Customer managed keys → Create key:
- Key type **Asymmetric**, usage **Sign and verify**, spec **ECC_NIST_P256**.
- Alias `alias/maxx-ap-seal`.
- Key administrators: `lenny-admin`. Key users: none for now (the Lambda role is added by policy below).

Key policy (replace `111122223333` with the account id; `MaxxApSealLambdaRole` is created in §4):

```json
{
  "Version": "2012-10-17",
  "Statement": [
    {
      "Sid": "AdministerButNeverSign",
      "Effect": "Allow",
      "Principal": {"AWS": "arn:aws:iam::111122223333:user/lenny-admin"},
      "Action": ["kms:Create*", "kms:Describe*", "kms:Enable*", "kms:List*", "kms:Put*", "kms:Update*",
                 "kms:Revoke*", "kms:Disable*", "kms:Get*", "kms:Delete*", "kms:TagResource", "kms:UntagResource",
                 "kms:ScheduleKeyDeletion", "kms:CancelKeyDeletion"],
      "Resource": "*"
    },
    {
      "Sid": "OnlyTheLambdaSigns",
      "Effect": "Allow",
      "Principal": {"AWS": "arn:aws:iam::111122223333:role/MaxxApSealLambdaRole"},
      "Action": ["kms:Sign", "kms:GetPublicKey", "kms:DescribeKey"],
      "Resource": "*"
    },
    {
      "Sid": "AnyoneMayReadThePublicKey",
      "Effect": "Allow",
      "Principal": {"AWS": "arn:aws:iam::111122223333:user/maxx-ap-portal"},
      "Action": ["kms:GetPublicKey"],
      "Resource": "*"
    }
  ]
}
```

Nobody has `kms:Sign` except the Lambda role; an asymmetric KMS key cannot be exported by anyone.
Turn on **automatic key rotation: off** (asymmetric keys do not rotate; a new key is a new `key_id`,
which the signature rows record).

## 3. DynamoDB tables

Create two on-demand tables (console → DynamoDB → Create table):

| Table | Partition key | Purpose |
| --- | --- | --- |
| `ap_passkeys` | `credential_id` (String, base64url) | Written at enrollment (`action: enroll`). Items: `user_id`, `public_key_b64`, `sign_count`, `label`, `status`, `enrolled_at`, `last_used_at`. |
| `ap_signing_sessions` | `user_id` (String) | Written at session start (`action: session_start`). Items: `credential_id`, `started_at`. |

Turn on **point-in-time recovery** on `ap_passkeys`. Enable deletion protection on both.

## 4. Lambda

Console → Lambda → Create function → Author from scratch:
- Name `maxx-ap-seal`, runtime **Python 3.12**, architecture arm64 is fine.
- Execution role: create new role `MaxxApSealLambdaRole`. After creation, replace its inline policy with:

```json
{
  "Version": "2012-10-17",
  "Statement": [
    {"Effect": "Allow", "Action": ["logs:CreateLogGroup", "logs:CreateLogStream", "logs:PutLogEvents"], "Resource": "arn:aws:logs:us-east-1:111122223333:*"},
    {"Effect": "Allow", "Action": ["kms:Sign", "kms:GetPublicKey", "kms:DescribeKey"], "Resource": "arn:aws:kms:us-east-1:111122223333:key/KEY-ID-HERE"},
    {"Effect": "Allow", "Action": ["dynamodb:GetItem", "dynamodb:PutItem", "dynamodb:UpdateItem"],
     "Resource": ["arn:aws:dynamodb:us-east-1:111122223333:table/ap_passkeys", "arn:aws:dynamodb:us-east-1:111122223333:table/ap_signing_sessions"]}
  ]
}
```

- Upload `handler.py` with its dependency. Build the zip on any machine:

```
mkdir build && pip install --target build "webauthn==2.5.0" "cbor2" "cryptography" --platform manylinux2014_aarch64 --only-binary=:all: --python-version 3.12
cp handler.py build/ && (cd build && zip -r ../seal.zip .)
```

  Handler: `handler.lambda_handler`. Timeout 10 s. Memory 256 MB.
- Environment variables: `KMS_KEY_ID=alias/maxx-ap-seal`, `RP_ID=accounting.90ten.life`,
  `EXPECTED_ORIGIN=https://accounting.90ten.life`, `PASSKEYS_TABLE=ap_passkeys`, `SESSIONS_TABLE=ap_signing_sessions`.

## 5. Function URL

Lambda → Configuration → Function URL → Create: **Auth type AWS_IAM**, no CORS. Copy the URL into the
VPS `.env` as `SIGNING_SERVICE_URL=`. Add a resource-based policy so only the portal user may invoke
(Lambda → Configuration → Permissions → Resource-based policy → Add → Function URL):

- Principal: `arn:aws:iam::111122223333:user/maxx-ap-portal`, Action `lambda:InvokeFunctionUrl`, Auth type `AWS_IAM`.

The portal signs every request with SigV4 (service `lambda`), so a stolen `.env` on the VPS gives an
attacker the ability to *ask* for seals — and every ask still needs a valid passkey assertion or a
recent session recorded by this function.

## 6. IAM policy for Andy's user `maxx-ap-portal` — exact JSON, nothing else

```json
{
  "Version": "2012-10-17",
  "Statement": [
    {
      "Sid": "CallTheSealService",
      "Effect": "Allow",
      "Action": "lambda:InvokeFunctionUrl",
      "Resource": "arn:aws:lambda:us-east-1:111122223333:function:maxx-ap-seal",
      "Condition": {"StringEquals": {"lambda:FunctionUrlAuthType": "AWS_IAM"}}
    },
    {
      "Sid": "WriteAuditAnchorsOnly",
      "Effect": "Allow",
      "Action": "s3:PutObject",
      "Resource": "arn:aws:s3:::maxx-ap-anchors/*"
    }
  ]
}
```

No `s3:GetObject`, no `s3:DeleteObject`, no `kms:*`, no `lambda:UpdateFunctionCode`, no DynamoDB.
Create one access key for this user and put it in the VPS `.env` as `AWS_ACCESS_KEY_ID` /
`AWS_SECRET_ACCESS_KEY`. Rotate it annually; Lenny can revoke it at any time without Andy.

## 7. Anchors bucket (for the audit chain, AUD-1)

S3 → Create bucket `maxx-ap-anchors`, **Object Lock enabled** at creation, default retention
**Compliance, 7 years**, block all public access, versioning on (required by Object Lock). The portal
writes `s3://maxx-ap-anchors/<date>/<ts>.json` every 15 minutes with the `s3:PutObject` permission above;
nobody — including Lenny — can shorten or delete those objects before the retention ends.

## 8. Checks before go-live

1. `aws kms get-public-key --key-id alias/maxx-ap-seal` as `maxx-ap-portal` works; `aws kms sign ...` as
   `maxx-ap-portal` is **denied**.
2. Enroll a test passkey in the portal; confirm an item appears in `ap_passkeys`.
3. Approve a test invoice; confirm `approval_signature.seal_key_id` is the KMS key ARN and that
   `accounting.approvals.signing_service.verify_seal(payload_sha256, seal, <public key PEM>)` returns True.
4. Stop the Lambda (set reserved concurrency to 0) and try to approve: the portal must show
   "Approvals are temporarily unavailable. Nothing was lost." and write no signature row.
5. Restore concurrency.

## What this does not do

It does not hold the invoice, the vendor or the amount — only the hash and the payload fields the
portal sends for recomputation. It does not know who the people are beyond the user ids. It never
calls back into the portal.

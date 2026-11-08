# WorkLog quarterly backup and retention

This workflow archives eligible completed handover records to an Excel workbook in Cloudinary.

## Retention policy

- Eligibility is based on the post's **creation date**: `createdAt <= now - 180 days`.
- Only records with `completed === true` are eligible.
- Pending records are never automatically deleted.
- The Excel workbook preserves post text and metadata. It records attachment metadata and URLs, but does **not** embed photo/video files.
- Photo/video originals are permanently deleted from Cloudinary only after the Excel upload has been verified.
- The Firestore document is deleted only after its associated Cloudinary media deletion steps complete.
- Missing attachment `publicId` values cause that record's deletion to stop rather than guessing a Cloudinary asset identifier.

## Schedule

The workflow is scheduled quarterly at 03:17 Korea Standard Time on January 1, April 1, July 1, and October 1. GitHub scheduled workflows can be delayed. It can also be started manually from the Actions tab.

## Required GitHub Actions secrets

Set these under **Settings → Secrets and variables → Actions → Repository secrets**:

- `FIREBASE_SERVICE_ACCOUNT_JSON`: Firebase Admin service-account JSON for project `proto-beta1`.
- `CLOUDINARY_CLOUD_NAME`: `worklog`.
- `CLOUDINARY_API_KEY`: Cloudinary API key.
- `CLOUDINARY_API_SECRET`: Cloudinary API secret. Never put this in `index.html`.

The Firebase service account should have only the Firestore permissions needed for this archive job. Keep all credentials in GitHub Actions secrets.

## Safe first run

1. Add the required secrets.
2. Open **Actions → WorkLog quarterly backup and retention → Run workflow**.
3. Leave `enable_delete` set to `false`. This creates and uploads the workbook but does not delete Firestore documents or Cloudinary media.
4. Download the workflow artifact `worklog-retention-report-...` and verify the candidate count and backup details.
5. Do not enable permanent deletion until the dry-run output and uploaded workbook have been reviewed.

For a later manual run, `enable_delete` must be explicitly set to `true` to permit permanent deletion. The repository variable `WORKLOG_ENABLE_DELETE` should remain unset or `false` unless an administrator deliberately chooses scheduled deletion.

## Important limitations

- Attachments are not embedded in the Excel file. Once originals are deleted, the recorded attachment URLs will no longer be usable.
- A backup run with no eligible records creates no workbook.
- Failed backup upload or verification blocks deletion.
- Review the generated run report after every run.

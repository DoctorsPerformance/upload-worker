# Upload worker

Files workbooks uploaded from the Doctors Performance dashboard into its (private) data
repository. Runs every 5 minutes on GitHub Actions (`.github/workflows/inbox.yml`).

- Uploads arrive encrypted; only this worker's key (a repository secret) can read them.
- Only the department folders listed in `inbox.js` can be written, and only Excel files with ID
  and Name columns are accepted.
- No data and no keys are stored in this repository.

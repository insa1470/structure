# Database Rollout Plan

The production app still defaults to JSON storage.

Current safe default:

```text
TASK_STORE=json
```

PostgreSQL can be enabled later with:

```text
TASK_STORE=postgres
DATABASE_URL=...
```

## Phase 1: Storage Adapter

`storage.py` now supports two task stores:

- `JsonTaskStore`: existing file-based storage under `app_data/tasks`.
- `PostgresTaskStore`: database-backed storage that preserves the full task JSON.

The first PostgreSQL schema intentionally keeps the complete task document in
`tasks.task` so existing API responses and frontend behavior do not change.

## Phase 2: Tables

The initial PostgreSQL adapter creates:

- `tasks`: one row per task, including status, error, summary, source files, and full task JSON.
- `task_files`: reserved table for uploaded file metadata.
- `task_snapshots`: append-only task snapshots for recovery and audit.

## Phase 3: Safe Cutover

Recommended sequence:

1. Keep production on `TASK_STORE=json`.
2. Deploy the adapter and confirm current JSON flows still work.
3. Create a staging environment with `TASK_STORE=postgres`.
4. Run upload, recognition, review, candidate, chart shareholder, draft, and print workflows.
5. Switch production to `TASK_STORE=postgres` only after staging passes.

Rollback is immediate: set `TASK_STORE=json` and redeploy.

## Later Normalization

After the PostgreSQL adapter is stable, split high-value fields into normalized
tables:

- `master_rows`
- `review_rows`
- `candidate_rows`
- `chart_shareholders`
- `task_events`
- `chart2_chunks`

Do this after the task store cutover, not before.

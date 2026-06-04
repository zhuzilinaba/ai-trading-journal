## Summary

-

## Key Changes

-

## How to Run Locally

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python trade_journal.py
```

## How to Test

```bash
python -m py_compile trade_journal.py
```

## Security Notes

- No API keys, `.env` files, exported workbooks, or real trading screenshots should be committed.
- Screenshots used in documentation must be sanitized.

## Screenshots

- [ ] Main app screenshot added or intentionally deferred
- [ ] Excel export screenshot added or intentionally deferred

## Checklist

- [ ] README is accurate
- [ ] App launches locally
- [ ] Excel export works
- [ ] No secrets are exposed
- [ ] Local data and generated files are ignored

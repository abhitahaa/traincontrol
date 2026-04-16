# TrainCycles Automation

Automates the CYCLES → testcycle pipeline, eliminating manual copy-paste.

## Files

| File | Purpose |
|------|---------|
| `run_pipeline.py` | **Run this** — auto-detects CYCLES file and runs both steps |
| `build_master.py` | Step 1: CYCLES → my_master.xlsx |
| `populate_testcycle.py` | Step 2: my_master.xlsx → testcycle_updated.xlsx |
| `testcycle.xlsx` | Template (do not modify) |

## How to Run

1. Drop your CYCLES `.xlsx` file into this folder
2. Run:
```bash
python run_pipeline.py
```

**Outputs:**
- `my_master.xlsx` — clean 6-sheet reference master
- `testcycle_updated.xlsx` — ready to upload to portal

## Requirements

```bash
pip install openpyxl pandas
```

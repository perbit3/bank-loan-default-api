# Bank Loan Default Prediction API

Random Forest classifier (500 trees) that estimates the probability that a bank customer will default on a loan, served as a REST API with FastAPI and called from a macro-enabled Excel workbook.

## Data

* `bank_loan_train.csv`: 700 customers used to train the model (BankLoan dataset)
* `bank_loan_test.csv`: 10 customers to score (BankLoan_Test dataset)

Target: `DEFAULTER` (1 = defaulted, 0 = did not default). Predictors: every other column except `SN` (`AGE`, `EMPLOY`, `ADDRESS`, `DEBTINC`, `CREDDEBT`, `OTHDEBT`).

## Files

| File | Purpose |
|---|---|
| `model.py` | Loads the data and trains the 500-tree Random Forest |
| `train.py` | Trains the model, prints out-of-bag accuracy and feature importance, scores the test dataset and saves `test_predictions.csv` |
| `main.py` | FastAPI app with the `/predict` endpoint |

## Run locally

```bash
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
python train.py
uvicorn main:app --reload
```

Open http://127.0.0.1:8000/docs for the interactive documentation.

## Endpoints

| Method | Path | Description |
|---|---|---|
| POST | `/predict` | JSON body with the six predictors, returns the probability of default |
| GET | `/predict` | Same, using query string parameters |
| GET | `/test-predictions` | Probability of default for every customer in the test dataset |
| GET | `/health` | Health check |

Example request:

```bash
curl -X POST https://bank-loan-default-api.onrender.com/predict \
  -H "Content-Type: application/json" \
  -d '{"AGE": 3, "EMPLOY": 16, "ADDRESS": 10, "DEBTINC": 9.1, "CREDDEBT": 11.3, "OTHDEBT": 5.01}'
```

Example response:

```json
{"probability_of_default": 0.68, "predicted_defaulter": 1, "risk_label": "Likely to default", "threshold": 0.5}
```

## Deployment

Deployed on Render (free plan) using `render.yaml`. Python is pinned to 3.11.9 with `.python-version`.

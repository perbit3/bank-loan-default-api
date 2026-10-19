"""FastAPI service that returns a customer's predicted probability of loan default."""
import pandas as pd
from fastapi import Depends, FastAPI
from fastapi.responses import RedirectResponse
from pydantic import BaseModel, Field

from model import FEATURES, N_TREES, TEST_FILE, load_data, predict_probability, train_model

app = FastAPI(
    title="Bank Loan Default Prediction API",
    description=(
        "Random Forest classifier (500 trees) trained on the BankLoan dataset. "
        "Send a customer's details to /predict to get the probability that the customer will default."
    ),
    version="1.0.0",
)

MODEL = train_model()  # trained once at start-up (fixed random_state, so results are reproducible)
THRESHOLD = 0.5


class Customer(BaseModel):
    AGE: int = Field(..., ge=1, le=3, description="Age group code (1, 2 or 3)")
    EMPLOY: float = Field(..., ge=0, description="Years with current employer")
    ADDRESS: float = Field(..., ge=0, description="Years at current address")
    DEBTINC: float = Field(..., ge=0, description="Debt to income ratio (%)")
    CREDDEBT: float = Field(..., ge=0, description="Credit card debt")
    OTHDEBT: float = Field(..., ge=0, description="Other debt")

    model_config = {
        "json_schema_extra": {
            "example": {"AGE": 3, "EMPLOY": 16, "ADDRESS": 10, "DEBTINC": 9.1, "CREDDEBT": 11.3, "OTHDEBT": 5.01}
        }
    }


class Prediction(BaseModel):
    probability_of_default: float
    predicted_defaulter: int
    risk_label: str
    threshold: float


def score(customer: Customer) -> Prediction:
    row = pd.DataFrame([customer.model_dump()], columns=FEATURES)
    prob = round(float(predict_probability(MODEL, row)[0]), 4)
    is_defaulter = int(prob >= THRESHOLD)
    return Prediction(
        probability_of_default=prob,
        predicted_defaulter=is_defaulter,
        risk_label="Likely to default" if is_defaulter else "Unlikely to default",
        threshold=THRESHOLD,
    )


@app.get("/", include_in_schema=False)
def root():
    return RedirectResponse(url="/docs")


@app.get("/health")
def health():
    return {"status": "ok", "model": "RandomForestClassifier", "trees": len(MODEL.estimators_)}


@app.post("/predict", response_model=Prediction)
def predict(customer: Customer):
    """Return the predicted probability of default for one customer (JSON body)."""
    return score(customer)


@app.get("/predict", response_model=Prediction)
def predict_get(customer: Customer = Depends()):
    """Same as POST /predict, using query string parameters (handy for testing in a browser)."""
    return score(customer)


@app.get("/test-predictions")
def test_predictions():
    """Probability of default for every customer in the BankLoan_Test dataset."""
    test_df = load_data(TEST_FILE)
    test_df["PROB_DEFAULT"] = predict_probability(MODEL, test_df).round(4)
    test_df["PREDICTED_DEFAULTER"] = (test_df["PROB_DEFAULT"] >= THRESHOLD).astype(int)
    return {"trees": N_TREES, "rows": test_df.to_dict(orient="records")}

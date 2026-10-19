"""Random Forest model for predicting the probability of loan default."""
from pathlib import Path

import pandas as pd
from sklearn.ensemble import RandomForestClassifier

BASE_DIR = Path(__file__).resolve().parent
TRAIN_FILE = BASE_DIR / "bank_loan_train.csv"
TEST_FILE = BASE_DIR / "bank_loan_test.csv"

TARGET = "DEFAULTER"
ID_COLUMN = "SN"
FEATURES = ["AGE", "EMPLOY", "ADDRESS", "DEBTINC", "CREDDEBT", "OTHDEBT"]
N_TREES = 500
RANDOM_STATE = 42


def load_data(path):
    return pd.read_csv(path)


def train_model(train_df=None):
    """Fit a 500-tree Random Forest using every column except SN as a predictor."""
    if train_df is None:
        train_df = load_data(TRAIN_FILE)
    X = train_df.drop(columns=[ID_COLUMN, TARGET])
    y = train_df[TARGET]
    model = RandomForestClassifier(
        n_estimators=N_TREES,
        oob_score=True,
        random_state=RANDOM_STATE,
        n_jobs=-1,
    )
    model.fit(X, y)
    return model


def predict_probability(model, df):
    """Return the probability of default (class 1) for each row of df."""
    return model.predict_proba(df[FEATURES])[:, 1]

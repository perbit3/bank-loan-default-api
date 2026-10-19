"""Train the Random Forest, report its performance and score the test dataset.

Run with:  python train.py
"""
import pandas as pd
from sklearn.metrics import accuracy_score, confusion_matrix, roc_auc_score

from model import (BASE_DIR, FEATURES, ID_COLUMN, TARGET, TEST_FILE, TRAIN_FILE,
                   load_data, predict_probability, train_model)


def main():
    # 1-4. Load the training data, use DEFAULTER as the target and all other
    #      variables except SN as predictors, and fit a 500-tree Random Forest.
    train_df = load_data(TRAIN_FILE)
    print(f"Training rows: {len(train_df)}  |  Predictors: {FEATURES}")
    print(f"Default rate in training data: {train_df[TARGET].mean():.2%}")

    model = train_model(train_df)
    print(f"Trees: {len(model.estimators_)}")
    print(f"Out-of-bag accuracy: {model.oob_score_:.4f}")

    importance = pd.Series(model.feature_importances_, index=FEATURES).sort_values(ascending=False)
    print("\nFeature importance:")
    print(importance.round(4).to_string())

    # 5. Load the test dataset and estimate the probability of default.
    test_df = load_data(TEST_FILE)
    test_df["PROB_DEFAULT"] = predict_probability(model, test_df).round(4)
    test_df["PREDICTED_DEFAULTER"] = (test_df["PROB_DEFAULT"] >= 0.5).astype(int)

    print("\nTest dataset predictions:")
    print(test_df[[ID_COLUMN] + FEATURES + [TARGET, "PROB_DEFAULT", "PREDICTED_DEFAULTER"]].to_string(index=False))

    y_true, y_prob = test_df[TARGET], test_df["PROB_DEFAULT"]
    print(f"\nTest accuracy (0.5 cut-off): {accuracy_score(y_true, test_df['PREDICTED_DEFAULTER']):.2%}")
    print(f"Test ROC AUC: {roc_auc_score(y_true, y_prob):.4f}")
    print("Confusion matrix [rows = actual 0/1, columns = predicted 0/1]:")
    print(confusion_matrix(y_true, test_df["PREDICTED_DEFAULTER"]))

    out = BASE_DIR / "test_predictions.csv"
    test_df.to_csv(out, index=False)
    print(f"\nSaved {out.name}")


if __name__ == "__main__":
    main()

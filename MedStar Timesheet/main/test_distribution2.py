import pandas as pd
from collections import defaultdict

def test_fair_distribution(csv_path):
    df = pd.read_csv(csv_path)

    # Remove NaN or empty values
    df["P_Assignment"] = df["P_Assignment"].astype(str).str.strip()
    df["W_Assignment"] = df["W_Assignment"].astype(str).str.strip()
    df = df[(df["P_Assignment"] != "") & (df["W_Assignment"] != "")]

    # Count assignments
    p_counts = defaultdict(int)
    w_counts = defaultdict(int)

    for name in df["P_Assignment"]:
        p_counts[name] += 1

    for name in df["W_Assignment"]:
        w_counts[name] += 1

    all_people = set(p_counts) | set(w_counts)

    # Test 1: Everyone got P and W at least once
    for person in all_people:
        assert p_counts[person] > 0, f"{person} never assigned P"
        assert w_counts[person] > 0, f"{person} never assigned W"

    # Test 2: P and W counts for each person differ by at most 1
    for person in all_people:
        diff = abs(p_counts[person] - w_counts[person])
        assert diff <= 1, f"{person} has imbalance (P: {p_counts[person]}, W: {w_counts[person]})"

    # Test 3: Total balance of P and W across entire schedule
    total_p = sum(p_counts.values())
    total_w = sum(w_counts.values())
    assert abs(total_p - total_w) <= 1, f"Total imbalance (P: {total_p}, W: {total_w})"

    print("✅ All fairness tests passed!")

# Run the test
if __name__ == "__main__":
    test_fair_distribution("precomputed_pw_assignments.csv")
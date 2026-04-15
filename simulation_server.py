#!/usr/bin/env python3
"""
ICS/OT EDR Simulation Server
Reads real SWaT CSV datasets and runs the EDRAnomalyDetector in real-time.
Streams sensor data + detection results to the dashboard.
"""
from flask import Flask, jsonify, request
from flask_cors import CORS
import pandas as pd
import numpy as np
import threading
import time
import os
from anomaly_detector import EDRAnomalyDetector

app = Flask(__name__)
CORS(app)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
SENSORS = [
    "FIT101","LIT101","AIT201","AIT202","AIT203","FIT201",
    "LIT301","FIT301","AIT401","FIT501","PIT501","PIT502","PIT503","FIT601"
]
ACTUATORS = [
    "P101","P102","P201","P203","P205","P301","P302",
    "P402","P403","P501","P602","UV401"
]
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SIM_SIZE = 3000          # rows per simulation run
NORMAL_LEAD = 500        # normal rows before attack begins (for training)

# ---------------------------------------------------------------------------
# Load & prepare datasets
# ---------------------------------------------------------------------------
print("\n" + "=" * 70)
print("  ICS/OT EDR — REAL-TIME SIMULATION SERVER")
print("=" * 70)

def load_csv(name, max_rows=None):
    path = os.path.join(BASE_DIR, name)
    if not os.path.exists(path):
        print(f"  [WARN] {name} not found")
        return None
    print(f"  Loading {name} …", end=" ", flush=True)
    df = pd.read_csv(path, nrows=max_rows)
    df.columns = [c.strip() for c in df.columns]
    if "Normal/Attack" in df.columns:
        df["Normal/Attack"] = df["Normal/Attack"].str.strip()
    print(f"{len(df)} rows")
    return df

normal_df = load_csv("normal.csv", max_rows=5000)
attack_df = load_csv("attack.csv", max_rows=5000)
merged_df = None  # built on demand

def build_sim_data(dataset_name):
    """Build a simulation sequence from the chosen dataset."""
    if dataset_name == "normal" and normal_df is not None:
        return normal_df.sample(n=min(SIM_SIZE, len(normal_df)), random_state=42).reset_index(drop=True)
    elif dataset_name == "attack" and attack_df is not None and normal_df is not None:
        # 500 normal rows (model training) then 2500 attack rows
        n_normal = min(NORMAL_LEAD, len(normal_df))
        n_attack = min(SIM_SIZE - n_normal, len(attack_df))
        part_n = normal_df.sample(n=n_normal, random_state=42)
        part_a = attack_df.sample(n=n_attack, random_state=42)
        return pd.concat([part_n, part_a], ignore_index=True)
    elif dataset_name == "merged":
        if normal_df is not None and attack_df is not None:
            half = SIM_SIZE // 2
            part_n = normal_df.sample(n=min(half, len(normal_df)), random_state=42)
            part_a = attack_df.sample(n=min(half, len(attack_df)), random_state=42)
            return pd.concat([part_n, part_a], ignore_index=True)
    # fallback
    if attack_df is not None:
        return attack_df.head(SIM_SIZE).copy()
    if normal_df is not None:
        return normal_df.head(SIM_SIZE).copy()
    return pd.DataFrame()

# ---------------------------------------------------------------------------
# Detector
# ---------------------------------------------------------------------------
detector = EDRAnomalyDetector(sensors=SENSORS, window_size=50, alpha=0.1, persistence_threshold=5)
print(f"  Detector ready  ({len(SENSORS)} sensors)")

# ---------------------------------------------------------------------------
# Global simulation state
# ---------------------------------------------------------------------------
sim = {
    "playing": False,
    "dataset": "attack",
    "speed": 5,            # rows per second
    "index": 0,
    "total": 0,
    "data": None,          # current DataFrame
    # live row state
    "sensors": {s: {"value": 0.0, "trend": 0, "alarm": False} for s in SENSORS},
    "actuators": {a: {"state": 0} for a in ACTUATORS},
    "is_attack": False,
    "ground_truth": "Normal",
    "risk_score": 0.0,
    "metrics": {},
    "alerts": [],
    "confusion": {"tp": 0, "fp": 0, "tn": 0, "fn": 0},
    "stats": {"total_processed": 0, "attacks_detected": 0},
}
prev_sensor_vals = {}

def reset_sim(dataset_name=None):
    global detector, prev_sensor_vals
    ds = dataset_name or sim["dataset"]
    sim["dataset"] = ds
    sim["data"] = build_sim_data(ds)
    sim["total"] = len(sim["data"])
    sim["index"] = 0
    sim["alerts"] = []
    sim["confusion"] = {"tp": 0, "fp": 0, "tn": 0, "fn": 0}
    sim["stats"] = {"total_processed": 0, "attacks_detected": 0}
    sim["risk_score"] = 0.0
    sim["is_attack"] = False
    sim["metrics"] = {}
    prev_sensor_vals = {}
    detector = EDRAnomalyDetector(sensors=SENSORS, window_size=50, alpha=0.1, persistence_threshold=5)
    print(f"  [SIM] Reset — dataset={ds}, rows={sim['total']}")

reset_sim("attack")

# ---------------------------------------------------------------------------
# Simulation thread
# ---------------------------------------------------------------------------
def process_row(row):
    """Process one CSV row through the detector."""
    global prev_sensor_vals

    # --- sensor values ---
    sensor_dict = {}
    for s in SENSORS:
        val = float(row.get(s, 0)) if pd.notna(row.get(s)) else 0.0
        prev = prev_sensor_vals.get(s, val)
        trend = 1 if val > prev else (-1 if val < prev else 0)
        sensor_dict[s] = {"value": val, "trend": trend, "alarm": False}
        prev_sensor_vals[s] = val
    sim["sensors"] = sensor_dict

    # --- actuator states (CSV uses 1=off, 2=on) ---
    for a in ACTUATORS:
        raw = int(row.get(a, 1)) if pd.notna(row.get(a)) else 1
        sim["actuators"][a] = {"state": 1 if raw == 2 else 0}

    # --- ground truth ---
    label = str(row.get("Normal/Attack", "Normal")).strip()
    is_attack = label == "Attack"
    sim["is_attack"] = is_attack
    sim["ground_truth"] = label

    # --- run detector ---
    det = detector.process(row.to_dict())
    risk   = det.get("risk_score", 0.0)
    is_anomaly = det.get("anomaly_detected", False)

    sim["risk_score"] = risk
    sim["metrics"] = {
        "D_M":        det.get("mahalanobis_distance", 0),
        "P_Xt":       det.get("probability_score", 1),
        "U_t":        det.get("unknown_score", 0),
        "H":          det.get("entropy", 0),
        "risk_score":  risk,
        "E_t":         det.get("reconstruction_error", 0),
        "T_score":     det.get("timing_score", 0),
        "is_anomaly":  is_anomaly,
        "anomaly_type": det.get("anomaly_type"),
        "severity":    det.get("severity", 0),
        "persistent":  det.get("persistent", False),
        "anomaly_flags": det.get("anomaly_flags", {}),
        # Isolation Forest metrics
        "iforest_score":   det.get("iforest_score", 0),
        "iforest_anomaly": det.get("iforest_anomaly", False),
        "iforest_trained": det.get("iforest_trained", False),
    }

    # --- mark alarms on the sensors that are furthest from baseline mean ---
    active_mu = detector.baseline_mu if detector.baseline_mu is not None else detector.mu
    active_sigma = detector.baseline_sigma if detector.baseline_sigma is not None else detector.sigma
    if is_anomaly and active_mu is not None:
        diffs = []
        for i, s in enumerate(SENSORS):
            v = sensor_dict[s]["value"]
            mu_val = active_mu[i]
            sigma_val = np.sqrt(active_sigma[i]) if active_sigma is not None else 1.0
            z = abs(v - mu_val) / (sigma_val + 1e-9)
            diffs.append((s, z))
        diffs.sort(key=lambda x: -x[1])
        for s, z in diffs[:3]:
            sim["sensors"][s]["alarm"] = True

    # --- confusion matrix ---
    sim["stats"]["total_processed"] += 1
    if is_anomaly:
        sim["stats"]["attacks_detected"] += 1
        if is_attack:
            sim["confusion"]["tp"] += 1
        else:
            sim["confusion"]["fp"] += 1
    else:
        if is_attack:
            sim["confusion"]["fn"] += 1
        else:
            sim["confusion"]["tn"] += 1

    # --- generate alert ---
    if is_anomaly:
        alert = {
            "id": len(sim["alerts"]) + 1,
            "ts": str(row.get("Timestamp", "")),
            "sensor": diffs[0][0] if diffs else "—",
            "type": det.get("anomaly_type", "UNKNOWN"),
            "sev": det.get("severity", 1),
            "rec": f"Risk {risk:.2%} | {'TRUE ATTACK' if is_attack else 'False Positive'}",
            "risk_score": risk,
            "method": det.get("anomaly_type", "Unknown"),
            "sample": sim["index"],
            "timestamp": str(row.get("Timestamp", "")),
        }
        sim["alerts"].append(alert)
        if len(sim["alerts"]) > 500:
            sim["alerts"] = sim["alerts"][-500:]


def simulation_loop():
    """Main simulation thread."""
    print("  [SIM] ▶ Playing …")
    while sim["playing"] and sim["index"] < sim["total"]:
        row = sim["data"].iloc[sim["index"]]
        process_row(row)
        sim["index"] += 1

        # progress log
        idx = sim["index"]
        if idx % 200 == 0 or idx == sim["total"]:
            tp = sim["confusion"]["tp"]
            fp = sim["confusion"]["fp"]
            fn = sim["confusion"]["fn"]
            tn = sim["confusion"]["tn"]
            print(f"    [{idx}/{sim['total']}]  TP={tp} FP={fp} TN={tn} FN={fn}  risk={sim['risk_score']:.3f}")

        time.sleep(1.0 / max(sim["speed"], 1))

    sim["playing"] = False
    print("  [SIM] ⏹ Stopped")


# ---------------------------------------------------------------------------
# API endpoints
# ---------------------------------------------------------------------------
@app.route("/api/state", methods=["GET"])
def get_state():
    cm = sim["confusion"]
    total_actual_attacks = cm["tp"] + cm["fn"]
    total_actual_normal  = cm["tn"] + cm["fp"]
    accuracy = ((cm["tp"] + cm["tn"]) / max(sim["stats"]["total_processed"], 1)) * 100

    return jsonify({
        # sensor / actuator state
        "sensors":    sim["sensors"],
        "actuators":  sim["actuators"],
        "is_attack":  sim["is_attack"],
        "risk_score": sim["risk_score"],
        "alerts":     sim["alerts"][-20:],
        # model metrics (from EDRAnomalyDetector)
        "metrics":    sim["metrics"],
        # simulation metadata
        "simulation": {
            "playing":      sim["playing"],
            "index":        sim["index"],
            "total":        sim["total"],
            "dataset":      sim["dataset"],
            "speed":        sim["speed"],
            "ground_truth": sim["ground_truth"],
            "detected":     sim["metrics"].get("is_anomaly", False),
        },
        # confusion matrix
        "confusion":  sim["confusion"],
        "accuracy":   round(accuracy, 1),
        # backwards-compat fields
        "attack_count": sim["stats"]["attacks_detected"],
        "normal_count": sim["stats"]["total_processed"] - sim["stats"]["attacks_detected"],
        "stats":        sim["stats"],
        "timestamp":    "",
    })


@app.route("/api/control", methods=["POST"])
def control():
    body = request.json or {}
    cmd  = body.get("command", "").lower()
    print(f"  [API] command={cmd}")

    if cmd in ("play", "start"):
        if not sim["playing"]:
            if sim["index"] >= sim["total"]:
                reset_sim()
            sim["playing"] = True
            threading.Thread(target=simulation_loop, daemon=True).start()
        return jsonify({"status": "playing"})

    elif cmd == "pause":
        sim["playing"] = False
        return jsonify({"status": "paused"})

    elif cmd == "reset":
        sim["playing"] = False
        time.sleep(0.15)
        reset_sim()
        return jsonify({"status": "reset"})

    elif cmd == "set_dataset":
        ds = body.get("dataset", "attack")
        sim["playing"] = False
        time.sleep(0.15)
        reset_sim(ds)
        return jsonify({"status": "dataset_changed", "dataset": ds})

    elif cmd == "set_speed":
        sim["speed"] = max(1, min(50, int(body.get("speed", 5))))
        return jsonify({"status": "speed_set", "speed": sim["speed"]})

    return jsonify({"status": "unknown_command"})


@app.route("/health")
def health():
    return jsonify({"status": "ok", "playing": sim["playing"]})


# ---------------------------------------------------------------------------
if __name__ == "__main__":
    print(f"\n  Server  →  http://localhost:5001")
    print(f"  Dashboard → open dashboard_v2.html (via simple_server on :3000)")
    print("=" * 70 + "\n")
    app.run(host="0.0.0.0", port=5001, debug=False, threaded=True)

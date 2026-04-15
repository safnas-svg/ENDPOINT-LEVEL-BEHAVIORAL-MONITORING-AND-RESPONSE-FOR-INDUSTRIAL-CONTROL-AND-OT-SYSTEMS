"""
Advanced ICS/OT EDR Anomaly Detection System
Implements all 12 mathematical formulas for statistical anomaly detection
+ Isolation Forest (ML-based) for hybrid ensemble detection
"""
import numpy as np
import pandas as pd
from collections import deque
from sklearn.ensemble import IsolationForest
import warnings
warnings.filterwarnings('ignore')

class EDRAnomalyDetector:
    """
    Implements the complete EDR formula set for real-time anomaly detection.
    Formulas from: Complete Mathematical Formulation of Advanced ICS/OT EDR System
    """
    
    def __init__(self, sensors, window_size=20, alpha=0.3, persistence_threshold=5):
        """
        Initialize the detector.
        
        Args:
            sensors: List of sensor names (e.g., ['FIT101', 'LIT101', ...])
            window_size: Sliding window size n (default 20)
            alpha: EWMA smoothing factor (default 0.3)
            persistence_threshold: Anomalies to confirm (default 5)
        """
        self.sensors = sensors
        self.m = len(sensors)  # Number of variables
        self.window_size = window_size
        self.alpha = alpha
        self.persistence_threshold = persistence_threshold
        
        # Weights for Risk Score and Unknown Behavior Score
        self.w = {'w1': 0.4, 'w2': 0.25, 'w3': 0.2, 'w4': 0.15}
        
        # Initialize state tracking
        self.window = deque(maxlen=window_size)  # Formula 2: W_t
        self.mu = None  # EWMA mean (Formula 3)
        self.sigma = None  # EWMA variance (Formula 4)
        self.covariance = None  # Covariance matrix Σ
        self.timing_deltas = deque(maxlen=50)  # Track timing deviations
        self.risk_history = deque(maxlen=10)  # Track risk scores for persistence
        self.prev_time = None
        self.prev_values = None
        
        # Normalization parameters
        self.sensor_ranges = {}
        self.value_history = {sensor: deque(maxlen=1000) for sensor in sensors}

        # Warmup: don't flag anomalies until baseline is established
        self.sample_count = 0
        self.warmup_samples = max(window_size * 4, 100)  # need enough data to learn baseline

        # Frozen baseline: lock the baseline after warmup so attacks can't shift it
        self.baseline_mu = None
        self.baseline_sigma = None
        self.baseline_cov_inv = None  # Frozen inverse covariance matrix

        # --- Isolation Forest (ML-based anomaly detection) ---
        self.iforest = IsolationForest(
            n_estimators=100,       # Number of isolation trees
            contamination=0.05,     # Expected proportion of anomalies (~5%)
            max_samples='auto',
            random_state=42
        )
        self.iforest_trained = False
        self.iforest_training_buffer = []  # Collect normal samples during warmup
    
    def update_sensor_ranges(self, df):
        """Learn sensor ranges from historical data."""
        for sensor in self.sensors:
            if sensor in df.columns:
                valid_values = pd.to_numeric(df[sensor], errors='coerce').dropna()
                if len(valid_values) > 0:
                    self.sensor_ranges[sensor] = (valid_values.min(), valid_values.max())
                else:
                    self.sensor_ranges[sensor] = (0, 1)
            else:
                self.sensor_ranges[sensor] = (0, 1)
    
    def _extract_state_vector(self, row):
        """
        Formula 1: Extract state vector X_t = [x1(t), x2(t), ..., xm(t)]^T
        
        Args:
            row: Dictionary with sensor values
            
        Returns:
            numpy array of shape (m,)
        """
        X_t = np.array([
            float(row.get(sensor, 0)) if pd.notna(row.get(sensor)) else 0
            for sensor in self.sensors
        ])
        return X_t
    
    def _update_window(self, X_t):
        """
        Formula 2: Update sliding window W_t = {X_{t-n+1}, ..., X_t}
        
        Args:
            X_t: Current state vector
        """
        self.window.append(X_t)
        self.value_history[self.sensors[0]].append(X_t[0])  # Track for timing
    
    def _update_ewma_mean(self, X_t):
        """
        Formula 3: Update EWMA mean μ_t = αX_t + (1-α)μ_{t-1}
        
        Args:
            X_t: Current state vector
            
        Returns:
            Updated mean vector μ_t
        """
        if self.mu is None:
            self.mu = X_t.copy()
        else:
            self.mu = self.alpha * X_t + (1 - self.alpha) * self.mu
        return self.mu
    
    def _update_ewma_variance(self, X_t):
        """
        Formula 4: Update EWMA variance σ_t^2 = α(X_t - μ_t)^2 + (1-α)σ_{t-1}^2
        
        Args:
            X_t: Current state vector
            
        Returns:
            Updated variance vector σ_t^2
        """
        diff = X_t - self.mu
        squared_diff = diff ** 2
        
        if self.sigma is None:
            self.sigma = squared_diff.copy()
        else:
            self.sigma = self.alpha * squared_diff + (1 - self.alpha) * self.sigma
        
        # Ensure minimum variance to avoid singularity
        self.sigma = np.maximum(self.sigma, 1e-6)
        return self.sigma
    
    def _compute_covariance_matrix(self):
        """
        Compute covariance matrix Σ from window.
        
        Returns:
            Covariance matrix of shape (m, m)
        """
        if len(self.window) < 2:
            # Return identity matrix scaled by variance
            return np.eye(self.m) * np.mean(self.sigma) if self.sigma is not None else np.eye(self.m)
        
        W = np.array(list(self.window))
        try:
            cov = np.cov(W.T)
            if cov.ndim == 0:  # Single variable case
                cov = np.array([[cov]])
            # Add regularization to ensure positive definiteness
            cov += np.eye(self.m) * 1e-6
            return cov
        except:
            return np.eye(self.m)
    
    def _compute_mahalanobis_distance(self, X_t):
        """
        Formula 5: Compute Mahalanobis distance D_M(t) = sqrt((X_t - μ_t)^T Σ^{-1} (X_t - μ_t))
        
        Args:
            X_t: Current state vector
            
        Returns:
            Mahalanobis distance (scalar)
        """
        if self.mu is None or self.sigma is None:
            return 0.0

        diff = X_t - self.mu

        try:
            # Use frozen covariance if available, otherwise compute live
            if self.baseline_cov_inv is not None:
                cov_inv = self.baseline_cov_inv
            else:
                self.covariance = self._compute_covariance_matrix()
                cov_inv = np.linalg.pinv(self.covariance)
            D_M = np.sqrt(np.dot(diff, np.dot(cov_inv, diff)))
        except:
            # Fallback: use standardized distance
            D_M = np.sqrt(np.sum((diff ** 2) / self.sigma))

        return float(D_M)
    
    def _compute_probability_score(self, D_M):
        """
        Formula 6: Compute probability score P(X_t) = exp(-0.5 * D_M_norm^2)

        Uses normalized D_M so the score varies meaningfully for SWaT-scale distances.

        Args:
            D_M: Mahalanobis distance

        Returns:
            Probability score in [0, 1]
        """
        # Normalize D_M to same scale used in risk score (SWaT normal D_M ~4, attack ~13+)
        D_M_norm = D_M / 10.0
        P = np.exp(-0.5 * D_M_norm ** 2)
        return float(np.clip(P, 0, 1))
    
    def _compute_reconstruction_error(self, X_t):
        """
        Formula 7: Compute reconstruction error E_t = ||X_t - μ_t||^2
        
        Args:
            X_t: Current state vector
            
        Returns:
            Reconstruction error (scalar)
        """
        if self.mu is None:
            return 0.0
        
        diff = X_t - self.mu
        E_t = float(np.sum(diff ** 2))
        return E_t
    
    def _compute_timing_deviation(self):
        """
        Formula 8: Compute timing deviation T_score = |Δt_i - μ_Δt| / σ_Δt
        
        Returns:
            Timing score (scalar)
        """
        if len(self.timing_deltas) < 2:
            return 0.0
        
        deltas = np.array(list(self.timing_deltas))
        mu_delta = np.mean(deltas)
        sigma_delta = np.std(deltas) + 1e-6  # Avoid division by zero
        
        # Current delta (using window size as proxy for timing)
        current_delta = len(self.window) / self.window_size
        
        T_score = abs(current_delta - mu_delta) / sigma_delta
        return float(np.clip(T_score, 0, 10))  # Clip to reasonable range
    
    def _compute_entropy(self):
        """
        Formula 9: Compute entropy H = - Σ P(s_i) log P(s_i)
        
        Returns:
            Entropy value (scalar)
        """
        if len(self.window) < 2:
            return 0.0
        
        W = np.array(list(self.window))
        # Discretize continuous data for entropy calculation
        bins = 10
        entropies = []
        
        for i in range(self.m):
            hist, _ = np.histogram(W[:, i], bins=bins, range=(W[:, i].min(), W[:, i].max() + 1e-6))
            probs = hist / hist.sum()
            probs = probs[probs > 0]  # Remove zero probabilities
            
            if len(probs) > 0:
                h = -np.sum(probs * np.log(probs + 1e-10))
                entropies.append(h)
        
        H = float(np.mean(entropies)) if entropies else 0.0
        return H
    
    def _compute_risk_score(self, D_M, anomaly_flags):
        """
        Formula 10: Compute risk score Score_t = w1*D_M + w2*A_f + w3*A_t + w4*A_s
        
        Args:
            D_M: Mahalanobis distance
            anomaly_flags: Dict with 'frequency', 'timing', 'sequence' indicators
            
        Returns:
            Risk score (scalar, normalized to [0, 1])
        """
        A_f = anomaly_flags.get('frequency', 0)  # Frequency anomaly
        A_t = anomaly_flags.get('timing', 0)      # Timing anomaly
        A_s = anomaly_flags.get('sequence', 0)    # Sequence/pattern anomaly
        
        # Normalize D_M to [0, 1] — tuned for SWaT (normal p95≈6, attack median≈13)
        norm_D_M = float(np.clip(D_M / 10, 0, 1))
        
        score = (
            self.w['w1'] * norm_D_M +
            self.w['w2'] * A_f +
            self.w['w3'] * A_t +
            self.w['w4'] * A_s
        )
        
        return float(np.clip(score, 0, 1))
    
    def _compute_unknown_behavior_score(self, P, E_t, T_score, H):
        """
        Formula 11: Compute unknown behavior U_t = w1(1-P) + w2*E_t + w3*T_score + w4*H
        
        Args:
            P: Probability score
            E_t: Reconstruction error
            T_score: Timing deviation score
            H: Entropy
            
        Returns:
            Unknown behavior score (normalized to [0, 1])
        """
        # Normalize components to [0, 1] — tuned for SWaT (E_t range ~20k-60k)
        norm_E_t = float(np.clip(E_t / 50000, 0, 1))
        norm_H = float(np.clip(H / 2.3, 0, 1))
        
        score = (
            self.w['w1'] * (1 - P) +
            self.w['w2'] * norm_E_t +
            self.w['w3'] * T_score +
            self.w['w4'] * norm_H
        )
        
        return float(np.clip(score, 0, 1))
    
    def _check_persistence(self, current_risk):
        """
        Formula 12: Check if anomaly persists Σ Risk_i > γ
        
        Args:
            current_risk: Current risk score
            
        Returns:
            Boolean indicating persistent anomaly
        """
        self.risk_history.append(current_risk)
        
        if len(self.risk_history) < self.persistence_threshold:
            return False
        
        # Sum of recent risks exceeds threshold
        recent_sum = np.sum(list(self.risk_history)[-self.persistence_threshold:])
        gamma = 1.2  # Persistence threshold — tuned for SWaT
        
        return recent_sum > gamma
    
    def _detect_anomaly_flags(self, X_t):
        """
        Detect specific anomaly types: frequency, timing, sequence.
        
        Args:
            X_t: Current state vector
            
        Returns:
            Dict with anomaly flags
        """
        flags = {
            'frequency': 0.0,
            'timing': 0.0,
            'sequence': 0.0
        }
        
        if self.prev_values is None:
            self.prev_values = X_t
            return flags
        
        # Frequency anomaly: large changes between consecutive samples
        delta = np.abs(X_t - self.prev_values)
        if np.mean(delta) > np.std(delta) * 3:
            flags['frequency'] = float(np.clip(np.mean(delta) / 10, 0, 1))
        
        # Timing anomaly: irregular update patterns
        if self.prev_time is not None:
            import time
            current_time = time.time()
            delta_t = current_time - self.prev_time
            expected_delta_t = 1.0  # Expected 1 second between updates
            
            if abs(delta_t - expected_delta_t) > expected_delta_t * 0.5:
                flags['timing'] = float(np.clip(abs(delta_t - expected_delta_t) / 5, 0, 1))
            
            self.prev_time = current_time
        else:
            import time
            self.prev_time = time.time()
        
        # Sequence anomaly: statistical deviation from normal patterns
        if self.mu is not None:
            distances = [np.linalg.norm(v - self.mu) for v in list(self.window)[-5:]]
            mean_dist = np.mean(distances)
            std_dist = np.std(distances)
            
            if mean_dist > std_dist * 2:
                flags['sequence'] = float(np.clip(mean_dist / 10, 0, 1))
        
        self.prev_values = X_t
        return flags
    
    def process(self, row):
        """
        Process a single row and compute all anomaly metrics.
        
        Args:
            row: Dictionary with sensor values
            
        Returns:
            Dict with anomaly detection results
        """
        self.sample_count += 1

        # Formula 1: Extract state vector
        X_t = self._extract_state_vector(row)

        # Formula 2: Update sliding window
        self._update_window(X_t)

        # Formula 3: Update EWMA mean
        mu_t = self._update_ewma_mean(X_t)

        # Formula 4: Update EWMA variance
        sigma_t = self._update_ewma_variance(X_t)

        # Collect training data for Isolation Forest during warmup
        if self.sample_count <= self.warmup_samples:
            self.iforest_training_buffer.append(X_t.copy())

        # Freeze baseline after warmup — so attacks can't shift the reference
        if self.sample_count == self.warmup_samples:
            self.baseline_mu = self.mu.copy()
            self.baseline_sigma = self.sigma.copy()
            # Freeze covariance matrix too
            cov = self._compute_covariance_matrix()
            try:
                self.baseline_cov_inv = np.linalg.pinv(cov)
            except:
                self.baseline_cov_inv = np.diag(1.0 / self.baseline_sigma)

            # Train Isolation Forest on normal baseline data
            training_data = np.array(self.iforest_training_buffer)
            self.iforest.fit(training_data)
            self.iforest_trained = True
            self.iforest_training_buffer = []  # Free memory

        # Use frozen baseline for distance computations once available
        active_mu = self.baseline_mu if self.baseline_mu is not None else self.mu
        active_sigma = self.baseline_sigma if self.baseline_sigma is not None else self.sigma

        # Formula 5: Compute Mahalanobis distance (against frozen baseline)
        # Temporarily swap mu/sigma to use baseline
        saved_mu, saved_sigma = self.mu, self.sigma
        self.mu, self.sigma = active_mu, active_sigma
        D_M = self._compute_mahalanobis_distance(X_t)
        self.mu, self.sigma = saved_mu, saved_sigma

        # Formula 6: Compute probability score
        P_Xt = self._compute_probability_score(D_M)

        # Formula 7: Compute reconstruction error (against baseline)
        if active_mu is not None:
            diff = X_t - active_mu
            E_t = float(np.sum(diff ** 2))
        else:
            E_t = 0.0

        # Formula 8: Compute timing deviation
        T_score = self._compute_timing_deviation()

        # Formula 9: Compute entropy
        H = self._compute_entropy()

        # Detect specific anomaly types
        anomaly_flags = self._detect_anomaly_flags(X_t)

        # Formula 10: Compute risk score
        risk_score = self._compute_risk_score(D_M, anomaly_flags)

        # Formula 11: Compute unknown behavior score
        unknown_score = self._compute_unknown_behavior_score(P_Xt, E_t, T_score, H)

        # Formula 12: Check persistence
        is_persistent = self._check_persistence(risk_score)

        # --- Isolation Forest scoring ---
        iforest_anomaly = False
        iforest_score = 0.0   # Anomaly score S(x) ∈ [0, 1], higher = more anomalous
        if self.iforest_trained:
            # score_samples returns negative scores: lower = more anomalous
            raw_score = self.iforest.score_samples(X_t.reshape(1, -1))[0]
            # Convert to anomaly score: S(x) = 2^(-E(h(x))/c(n))
            # sklearn returns the negative of the offset-adjusted score
            # We convert: closer to 1 = anomaly, closer to 0 = normal
            iforest_score = float(np.clip(0.5 - raw_score, 0, 1))
            # Isolation Forest prediction: -1 = anomaly, 1 = normal
            iforest_pred = self.iforest.predict(X_t.reshape(1, -1))[0]
            iforest_anomaly = (iforest_pred == -1)

        # Determine anomaly type and severity
        anomaly_detected = False
        anomaly_type = None
        severity = 0

        # Don't flag during warmup
        if self.sample_count > self.warmup_samples:
            # --- HYBRID ENSEMBLE DECISION ---
            # Statistical model detection (D_M thresholds)
            stat_anomaly = False
            stat_severity = 0
            stat_type = None

            if D_M > 15:
                stat_anomaly = True
                stat_type = "IDENTITY"
                stat_severity = 3
            elif D_M > 8:
                stat_anomaly = True
                stat_type = "FREQUENCY" if anomaly_flags['frequency'] > 0.3 else "TIMING"
                stat_severity = 2
            elif D_M > 5:
                stat_anomaly = True
                stat_type = anomaly_flags['frequency'] > anomaly_flags['timing'] and anomaly_flags['frequency'] > 0 and "FREQUENCY" or "TIMING"
                stat_severity = 1

            # Hybrid decision: either model can flag, but both agreeing boosts severity
            if stat_anomaly and iforest_anomaly:
                # Both models agree → high confidence, boost severity
                anomaly_detected = True
                anomaly_type = stat_type
                severity = min(stat_severity + 1, 3)
            elif stat_anomaly:
                # Only statistical model flags → trust it (primary detector)
                anomaly_detected = True
                anomaly_type = stat_type
                severity = stat_severity
            elif iforest_anomaly and iforest_score > 0.6:
                # Only Isolation Forest flags with high confidence →
                # catches non-linear patterns the statistical model misses
                anomaly_detected = True
                anomaly_type = "PATTERN"
                severity = 2 if iforest_score > 0.75 else 1

        return {
            'anomaly_detected': anomaly_detected,
            'anomaly_type': anomaly_type,
            'severity': severity,
            'risk_score': risk_score,
            'unknown_score': unknown_score,
            'mahalanobis_distance': D_M,
            'probability_score': P_Xt,
            'reconstruction_error': E_t,
            'timing_score': T_score,
            'entropy': H,
            'persistent': is_persistent,
            'anomaly_flags': anomaly_flags,
            # Isolation Forest metrics
            'iforest_score': iforest_score,
            'iforest_anomaly': iforest_anomaly,
            'iforest_trained': self.iforest_trained,
        }

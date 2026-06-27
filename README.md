# Ruler vs Ruled: Monte Carlo Simulation and Analysis Pipeline

Created by Larzsolice (https://medium.com/@larzsolice or https://substack.com/@thepragmaticrealist).

This repository contains the complete, high-performance simulation engine and analytical pipeline for the Ruler vs Ruled system dynamics model. The project models the interaction between a governing entity (Ruler) and a population (Ruled) as an escalating, multi-turn game of Chicken with a dynamic, wasting resource pool.

## Overview

In this environment, societal perceived value (V_res) acts as a shared resource pool representing general prosperity. In each step of a simulation, both actors independently choose to act Hawkish (escalatory) or Dovish (accommodating).

Hawkish actions extract higher costs from the societal pool.

Dovish actions add value back (prosperity preservation), represented as negative costs.

Early state failure (Revolution or Crackdown) is triggered when collective costs exhaust the total available value (V_total = 10.0).


## Simulation Frameworks

The pipeline runs eight distinct simulation frameworks to isolate the impact of strategic baselines, starting configurations, and environmental reactivity:

* Simulation 1 (Stochastic Baseline): Completely random probability allocations at every step to map pure systemic noise.
* Simulation 2 (Blind Control): Fixed ideological quadrants (DD, DH, HD, HH, RR) oblivious to resource decay.
* Simulation 3 (Neutral Players): Random base hawkishness [0, 1] penalized by conscious awareness of structural decay (Aversion Penalties).
* Simulation 4 (Neutral TFT): Players copy the opponent's previous move reactively (Hawkishness is restricted to [0, 0.5) if the opponent played Dove, and (0.5, 1.0] if the opponent played Hawk).
* Simulation 5 (Aggressive Ruler): Ruler is hard-coded to be aggressive (P_Hawk > 0.5 before penalties); Ruled is neutral [0, 1].
* Simulation 6 (Aggressive Ruler TFT): Ruler is aggressive (>0.5); Ruled adopts Tit-for-Tat copying.
* Simulation 7 (Aggressive Ruled): Ruler is neutral; Ruled is hard-coded to be aggressive (>0.5).
* Simulation 8 (Aggressive Ruled TFT): Ruler adopts TFT copying; Ruled remains aggressive (>0.5).


## Strategic National Archetypes (Sims 3–8)

Societies are modeled with five explicit aversion penalties representing how they de-escalate under environmental stress:

* Equals: Symmetric linear de-escalation as resource pools drop.
* Cowards: Symmetric quadratic de-escalation (highly risk-averse).
* Fools: Symmetric sub-linear (square-root) de-escalation (highly resistant to compromise).
* Tyrants: Asymmetric de-escalation (Ruler is insulated from decay; Ruled de-escalates linearly).
* Brinksmen: Asymmetric de-escalation (Ruled is insulated from decay; Ruler de-escalates linearly).


## Analytical Methodology & Data Processing

The pipeline automates several advanced data science and clustering operations:


### 1. Principal Component Analysis (PCA)

To visualize the multi-dimensional behavior space in a 2D coordinate map, PCA compresses six core behavioral metrics:
* Average Ruler and Ruled Hawkishness (P_Hawk)
* Average Ruler and Ruled Costs incurred
* Average Ruler and Ruled Aversion Penalties applied

To prevent scale-drowning and ensure that spatial clustering represents genuine behavioral manifolds, the Prosperity Score (V_res) is explicitly withheld from the input PCA matrix.


### 2. Dynamically Optimized KDE Valley Clustering

Rather than using arbitrary thresholds, the pipeline segments outcome prosperity curves using Kernel Density Estimation (KDE) on Cumulative V_res:
* Sweeps bandwidth factors at fine resolution (0.01) on the persistent multiprocessing workers.
* Dynamically filters out the top 5 highest peak-producing bandwidths to prevent high-frequency noise and over-fitting.
* Discovers local density minima ("valleys") from the remaining stable fits to mathematically partition outcomes into discrete, ordered performance clusters.


### 3. High-Performance Multiprocessing Architecture

The engine initializes a persistent daemon pool of 10 worker processes on startup. Tasks (Simulations, KDE bandwidth sweeps, and PCA plots) are routed through process-safe queues (multiprocessing.Queue). This completely bypasses the massive overhead of repeatedly spawning and destroying process pools.


### Visual Marker Conventions (PCA Plots)

The generated plots use precise marker styles to immediately distinguish elite trajectories from failed societies:
* Standard Survivals: Unbordered, semi-transparent colored circles (o).
* Failed Societies: Solid colored crosses (x) corresponding to their KDE cluster.
* Top 100 Elite Performers: Marked with black-bordered semi-transparent diamonds (d).
* Bottom 100 Lowest Performers: Marked with solid black crosses if they failed, or empty black circles (o) if they survived despite degraded conditions.


## Output Files

Upon successful execution, all generated assets are exported and compressed directly into a ZIP archive containing:
* Report.docx: A professional, fully typeset Microsoft Word report using Candara typography, left-aligned headings, justified paragraphs, centered figures, and archetype-specific state failure tables.
* results_summary.xlsx: Complete outcome metrics per simulation run.
* results_pca_features.xlsx: Multi-dimensional feature matrix mapped to the calculated global PCA coordinates.
* params.txt: A reference sheet outlining all configuration parameters and the dynamically selected KDE bandwidth.
* Graphs/: A subfolder housing the isolated local and global behavioral PCA plots, count abundance sweeps, multi-panel early failure ratios, and 4-panel performance subplots.


## Requirements

Ensure you have the following third-party Python packages installed:

    pip install pandas numpy scipy matplotlib seaborn scikit-learn python-docx openpyxl



** Configuration & Parameter Adjustments##

At the absolute top of the ruler_ruled_sim.py file, you will find the main simulation configuration block designed for rapid adjustment:

    # ==========================================
    # SIMULATION PARAMETERS 
    # ==========================================
    PERCEIVED_VALUE = 10.0  # Value of a society
    MAX_COST_PER_STEP = 1.0 # Per player
    MAX_STEPS = 1000        # Per simulation
    N_REPS = 1000           # Number of repetitions for each simulated situation
    RECORD_HISTORY = False  # Slow and resource intensive, use on smaller N_REPS only


## Calibration Recommendations

Adjusting Prosperity vs. Step Friction: When modifying the simulation's baseline lifespan or difficulty, strongly prefer adjusting PERCEIVED_VALUE rather than editing MAX_COST_PER_STEP.

The Mathematical Reason: The strategic national archetypes rely on fractional aversion decay curves calibrated to $\frac{V_{\text{res}}}{V_{\text{total}}}$. Altering MAX_COST_PER_STEP shifts step cost distributions and unevenly distorts the de-escalation thresholds, whereas adjusting PERCEIVED_VALUE scales the entire system dynamics model cleanly.

## Running the Pipeline

Run the orchestrator locally from your terminal. Progressive progress bars track both the module imports, the parallel simulation batches, the dynamic KDE optimization, and the high-performance PCA renderings:

    python ruler_ruled_sim.py


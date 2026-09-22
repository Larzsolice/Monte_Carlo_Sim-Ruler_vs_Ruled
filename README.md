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
* Simulation 8 (Blind Control TFT): Same as Blind Control, but both players play Tit-for-Tat.


## Strategic National Archetypes (Sims 3–8)

Societies are modeled with five explicit aversion penalties representing how they de-escalate under environmental stress:

* Equals: Symmetric linear de-escalation as resource pools drop.
* Cowards: Symmetric quadratic de-escalation (highly risk-averse).
* Fools: Symmetric sub-linear (square-root) de-escalation (highly resistant to compromise).
* Tyrants: Asymmetric de-escalation (Ruler is insulated from decay; Ruled de-escalates linearly).
* Brinksmen: Asymmetric de-escalation (Ruled is insulated from decay; Ruler de-escalates linearly).


## Graphical Representations

The Prosperity Scores and Termination Proportions are displayed in paired bar graphs to assess how well societies perform.

The Prosperity Score is the sum of Residual Perceived Value (after Net Costs are subtracted) at the end of each step of the simulation.

The Termination Proportions display how many simulations ended with state failure, classifying termination events as either Revolution or government Crackdown.


## Requirements

Ensure you have the following third-party Python packages installed:

    pip install pandas numpy scipy matplotlib seaborn scikit-learn python-docx 

The simulations are faster if you have a CUDA-enabled version of pytorch installed.



** Configuration & Parameter Adjustments##

At the absolute top of the ruler_ruled_sim.py file, you will find the main simulation configuration block designed for rapid adjustment:

    # ==========================================
    # SIMULATION PARAMETERS 
    # ==========================================
    PERCEIVED_VALUE = 8.0   # Value of a society
    MAX_COST_PER_STEP = 1.0 # Per player
    MAX_STEPS = 1000        # Per simulation
    N_REPS = 1000000        # Number of repetitions for each simulated situation


## Calibration Recommendations

Adjusting Prosperity vs. Step Friction: When modifying the simulation's baseline lifespan or difficulty, strongly prefer adjusting PERCEIVED_VALUE rather than editing MAX_COST_PER_STEP.

The Mathematical Reason: The strategic national archetypes rely on fractional aversion decay curves calibrated to $\frac{V_{\text{res}}}{V_{\text{total}}}$. Altering MAX_COST_PER_STEP shifts step cost distributions and unevenly distorts the de-escalation thresholds, whereas adjusting PERCEIVED_VALUE scales the entire system dynamics model cleanly.

## Running the Pipeline

Run the orchestrator locally from your terminal. Progressive progress bars track both the module imports, the parallel simulation batches, the dynamic KDE optimization, and the high-performance PCA renderings:

    python ruler_ruled_sim.py


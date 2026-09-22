"""
Ruler vs Ruled - Monte Carlo Simulation and Analysis Pipeline
Created by the Pragmatic Realist
Find me on Substack: https://thepragmaticrealist.substack.com
Coded with assistance from Gemini

-------------------------------------------------------------

Required Third-Party Libraries:
- pandas (Data manipulation and Excel export)
- numpy (Numerical operations)
- scipy (Statistical operations)
- matplotlib & seaborn (Data visualization)
- python-docx (Automated Word report generation)
- openpyxl (Excel file generation, utilized by pandas under the hood)
- pytorch (Must be a CUDA version for GPU access)

Simulation Overview:
- Simulation 1: Stochastic Baseline.
- Simulation 2: Blind Control mapped over fixed starting Quadrants (DD, DH, HD, HH, RR).
- Simulation 3 & 4: Neutral Players (Base) vs Neutral Players (Tit-For-Tat) across 5 Archetypes.
- Simulation 5 & 6: Aggressive Ruler (Base) vs Aggressive Ruler (Tit-For-Tat).
- Simulation 7 & 8: Aggressive Ruled (Base) vs Aggressive Ruled (Tit-For-Tat).
- Simulation 9: Blind Control mapped over fixed starting Quadrants, utilizing Tit-For-Tat.

For N_reps = 1000000, you will need 76 GB of disc space if RECORD_HISTORY is False. This can be almost halved if files are not repackaged into a zip file.
"""

if __name__ == "__main__": print("\nInitializing simulation environment...", flush=True)

# ==========================================
# SIMULATION PARAMETERS 
# ==========================================
PERCEIVED_VALUE = 8.0   # Value of a society
MAX_COST_PER_STEP = 1.0 # Per player
MAX_STEPS = 1000        # Per simulation
N_REPS = 10000          # Number of repetitions for each simulated situation
RECORD_HISTORY = False  # Slow and resource intensive, use on smaller N_REPS only
SORT_DATA = False       # If True, large dataset exports will be sorted before archiving (Requires Memory)

TEMP_DIR = "temp"
OUTPUT_PREFIX = "Ruler_Vs_Ruled_Simulation"
FILE_REPORT = "01_Simulation_Report.docx"
FILE_SUMMARY = "02_Simulation_Summary_Results.csv"
FILE_FEATURES = "03_Simulation_Features_Results.csv"
FILE_HISTORY = "04_Simulation_Turn_History.csv"
FILE_PARAMS = "05_Simulation_Parameters.txt"

STANDARD_ARCHETYPE_ORDER = ['Equals', 'Cowards', 'Fools', 'Brinksmen', 'Tyrants', 'Control']
SIM_LABELS = {
    1: "Sim 1 Base", 2: "Sim 2 Control Base", 9: "Sim 9 Control TFT", 
    3: "Sim 3 Neutral Base", 4: "Sim 4 Neutral TFT", 
    5: "Sim 5 Agg Ruler Base", 6: "Sim 6 Agg Ruler TFT", 
    7: "Sim 7 Agg Ruled Base", 8: "Sim 8 Agg Ruled TFT"
}
SIM_PAIRS = [
    (2, 9, 'Control (Sim 2 vs 9)'),
    (3, 4, 'Neutral (Sim 3 vs 4)'),
    (5, 6, 'Aggressive Ruler (Sim 5 vs 6)'),
    (7, 8, 'Aggressive Ruled (Sim 7 vs 8)')
]

# ==========================================
# SLOW IMPORT LOADING
# ==========================================
n_import_chunks = 6

def print_import_progress(step, name):
    bar_length = 30
    filled = int(bar_length * step / n_import_chunks)
    bar = '#' * filled + '-' * (bar_length - filled)
    percent = (step / n_import_chunks) * 100
    if name == "":
        print(f"\rLoading core libraries: [{bar}] {percent:.0f}%                 \n", end="", flush=True)
    else:
        print(f"\rLoading core libraries: [{bar}] {percent:.0f}% ({name})     ", end="", flush=True)

# Fast imports
if __name__ == "__main__": print_import_progress(1, "Fast Imports")
import os
import shutil
import random
import zipfile
import multiprocessing
import time
import queue

# Workaround for OpenMP runtime conflict (OMP: Error #15)
os.environ['KMP_DUPLICATE_LIB_OK'] = 'TRUE'



# Slow imports
if __name__ == "__main__": print_import_progress(2, "pandas/numpy/scipy")
import pandas as pd
import numpy as np
import scipy.stats as stats
import warnings


if __name__ == "__main__": print_import_progress(3, "matplotlib")
import matplotlib.pyplot as plt
from matplotlib.patches import Patch

if __name__ == "__main__": print_import_progress(4, "seaborn")
import seaborn as sns

if __name__ == "__main__": print_import_progress(5, "python-docx")
import docx
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import RGBColor


if __name__ == "__main__": print_import_progress(6, "pytorch")
HAS_TORCH = False
HAS_CUDA = False
try:
    # pass
    import torch
    HAS_TORCH = True
    HAS_CUDA = torch.cuda.is_available()
    # print(HAS_CUDA)
    if HAS_CUDA:
        # Enable CUDNN benchmarking for faster tensor execution
        torch.backends.cudnn.benchmark = True
except ImportError:
    HAS_TORCH = False
    HAS_CUDA = False

if HAS_TORCH and int(np.__version__.split('.')[0]) >= 2:
    try:
        with warnings.catch_warnings(record=True) as w:
            warnings.simplefilter("always")
            # This dummy conversion triggers the NumPy C-API warning if incompatible
            _ = torch.zeros(1).numpy() 
            if len(w) > 0:
                for warn in w:
                    if "Failed to initialize NumPy" in str(warn.message):
                        if __name__ == "__main__": 
                            print(f"\n\n[WARNING] PyTorch is incompatible with NumPy {np.__version__} (PyTorch expects NumPy 1.x).")
                            print("[WARNING] GPU Acceleration has been DISABLED to prevent crashes.")
                            print("[WARNING] To fix this and re-enable the GPU, please run the following command in your terminal:")
                            print('          pip install "numpy<2"\n')
                        HAS_CUDA = False
                        HAS_TORCH = False
                        break
    except Exception:
        HAS_CUDA = False
        HAS_TORCH = False

if __name__ == "__main__": 
    print_import_progress(n_import_chunks, "")

# ==========================================
# 3. CORE SIMULATION ENGINE & GPU ACCELERATION
# ==========================================

def csv_writer_process(queue, summary_filepath, features_filepath):
    import pandas as pd
    feat_cols = ['Sim_Type', 'Case', 'Archetype', 'Terminal_Status', 'Avg_P_R', 'Avg_P_D', 'Avg_C_R', 'Avg_C_D', 'Avg_Pen_R', 'Avg_Pen_D', 'Cum_V_res']
    while True:
        msg = queue.get()
        if msg is None:
            break
        try:
            msg.to_csv(summary_filepath, mode='a', header=False, index=False)
            msg[feat_cols].to_csv(features_filepath, mode='a', header=False, index=False)
        except Exception:
            pass

def run_games_cuda(sim_type, case=None, archetype=None, n_games=10000, queue=None):
    """
    Simulate n_games in parallel on GPU using PyTorch vectorized CUDA tensors.
    Scales seamlessly to millions of simulations with minimal memory overhead.
    """
    device = torch.device("cuda" if HAS_CUDA else "cpu")

    C_R = torch.zeros(n_games, device=device, dtype=torch.float32)
    C_D = torch.zeros(n_games, device=device, dtype=torch.float32)
    cum_v_res = torch.zeros(n_games, device=device, dtype=torch.float32)
    step = torch.ones(n_games, device=device, dtype=torch.int32)
    active = torch.ones(n_games, device=device, dtype=torch.bool)

    # Initial random parameters
    if case == 'DD':
        case_p_R = torch.rand(n_games, device=device) * 0.5
        case_p_D = torch.rand(n_games, device=device) * 0.5
    elif case == 'DH':
        case_p_R = torch.rand(n_games, device=device) * 0.5
        case_p_D = 0.5 + torch.rand(n_games, device=device) * 0.5
    elif case == 'HD':
        case_p_R = 0.5 + torch.rand(n_games, device=device) * 0.5
        case_p_D = torch.rand(n_games, device=device) * 0.5
    elif case == 'HH':
        case_p_R = 0.5 + torch.rand(n_games, device=device) * 0.5
        case_p_D = 0.5 + torch.rand(n_games, device=device) * 0.5
    else:  # 'RR' or None
        case_p_R = torch.rand(n_games, device=device)
        case_p_D = torch.rand(n_games, device=device)

    p_R_base, p_D_base = case_p_R.clone(), case_p_D.clone()

    if sim_type in [3, 4]:
        p_R_base = torch.rand(n_games, device=device)
        p_D_base = torch.rand(n_games, device=device)
    elif sim_type in [5, 6]:
        p_R_base = 0.50001 + torch.rand(n_games, device=device) * 0.49999
        p_D_base = torch.rand(n_games, device=device)
    elif sim_type in [7, 8]:
        p_R_base = torch.rand(n_games, device=device)
        p_D_base = 0.50001 + torch.rand(n_games, device=device) * 0.49999

    p_R_hawk_if_H = 0.50001 + torch.rand(n_games, device=device) * 0.49999
    p_R_hawk_if_D = torch.rand(n_games, device=device) * 0.49999
    p_D_hawk_if_H = 0.50001 + torch.rand(n_games, device=device) * 0.49999
    p_D_hawk_if_D = torch.rand(n_games, device=device) * 0.49999

    # Running totals for means
    sum_p_R = torch.ones(n_games, device=device)
    sum_p_D = torch.zeros(n_games, device=device)
    sum_pen_R = torch.ones(n_games, device=device)
    sum_pen_D = torch.ones(n_games, device=device)

    # Step 1 execution
    dc_R = torch.rand(n_games, device=device) * MAX_COST_PER_STEP
    dc_D = torch.rand(n_games, device=device) * MAX_COST_PER_STEP
    contrib_R = dc_R
    contrib_D = -dc_D
    delta_C = (contrib_R + contrib_D) / 2.0

    C_R = torch.clamp(C_R + delta_C, min=0.0)
    C_D = torch.clamp(C_D + delta_C, min=0.0)
    v_res = torch.clamp(PERCEIVED_VALUE - (C_R + C_D), min=0.0)
    cum_v_res += v_res

    sum_c_R = contrib_R / 2.0
    sum_c_D = contrib_D / 2.0

    prev_act_R = torch.ones(n_games, device=device, dtype=torch.bool) # 'H'
    prev_act_D = torch.zeros(n_games, device=device, dtype=torch.bool) # 'D'

    last_contrib_R = contrib_R.clone()
    last_contrib_D = contrib_D.clone()

    current_step = 1
    while active.any() and current_step < MAX_STEPS:
        current_step += 1
        v_res = torch.clamp(PERCEIVED_VALUE - (C_R + C_D), min=0.0)
        
        # Calculate environmental decay penalties
        if sim_type >= 3 and sim_type != 9:
            penalty = v_res / PERCEIVED_VALUE
            if archetype == 'Equals':
                mul_R, mul_D = penalty, penalty
            elif archetype == 'Cowards':
                mul_R, mul_D = penalty ** 2, penalty ** 2
            elif archetype == 'Fools':
                mul_R, mul_D = penalty ** 0.5, penalty ** 0.5
            elif archetype == 'Brinksmen':
                mul_R, mul_D = penalty, 1.0 - (1.0 - penalty) / 3.0
            elif archetype == 'Tyrants':
                mul_R, mul_D = 1.0 - (1.0 - penalty) / 3.0, penalty
            else:
                mul_R, mul_D = torch.ones_like(penalty), torch.ones_like(penalty)
        else:
            mul_R, mul_D = torch.ones_like(v_res), torch.ones_like(v_res)

        # Calculate current turn probabilities
        if sim_type == 1:
            p_R_curr = torch.rand(n_games, device=device)
            p_D_curr = torch.rand(n_games, device=device)
        elif sim_type == 2:
            p_R_curr = p_R_base.clone()
            p_D_curr = p_D_base.clone()
        else:
            if current_step == 2:
                base_R, base_D = p_R_base, p_D_base
            else:
                if sim_type in [3, 5, 7]:
                    base_R, base_D = p_R_base, p_D_base
                elif sim_type in [4, 9]:
                    base_R = torch.where(prev_act_D, p_R_hawk_if_H, p_R_hawk_if_D)
                    base_D = torch.where(prev_act_R, p_D_hawk_if_H, p_D_hawk_if_D)
                elif sim_type == 6:
                    base_R = p_R_base
                    base_D = torch.where(prev_act_R, p_D_hawk_if_H, p_D_hawk_if_D)
                elif sim_type == 8:
                    base_R = torch.where(prev_act_D, p_R_hawk_if_H, p_R_hawk_if_D)
                    base_D = p_D_base

            p_R_curr = torch.clamp(base_R * mul_R, min=0.01)
            p_D_curr = torch.clamp(base_D * mul_D, min=0.01)

        # Action resolution
        act_R = torch.rand(n_games, device=device) < p_R_curr
        act_D = torch.rand(n_games, device=device) < p_D_curr

        dc_R = torch.rand(n_games, device=device) * MAX_COST_PER_STEP
        dc_D = torch.rand(n_games, device=device) * MAX_COST_PER_STEP

        contrib_R = torch.where(act_R, dc_R, -dc_R)
        contrib_D = torch.where(act_D, dc_D, -dc_D)
        delta_C = (contrib_R + contrib_D) / 2.0

        # Accumulate metrics only for active games
        C_R += torch.where(active, delta_C, 0.0)
        C_R = torch.clamp(C_R, min=0.0)
        C_D += torch.where(active, delta_C, 0.0)
        C_D = torch.clamp(C_D, min=0.0)

        v_res = torch.clamp(PERCEIVED_VALUE - (C_R + C_D), min=0.0)
        cum_v_res += torch.where(active, v_res, 0.0)

        sum_p_R += torch.where(active, p_R_curr, 0.0)
        sum_p_D += torch.where(active, p_D_curr, 0.0)
        sum_pen_R += torch.where(active, mul_R, 0.0)
        sum_pen_D += torch.where(active, mul_D, 0.0)
        sum_c_R += torch.where(active, contrib_R / 2.0, 0.0)
        sum_c_D += torch.where(active, contrib_D / 2.0, 0.0)

        last_contrib_R = torch.where(active, contrib_R, last_contrib_R)
        last_contrib_D = torch.where(active, contrib_D, last_contrib_D)

        step += torch.where(active, 1, 0)
        prev_act_R, prev_act_D = act_R, act_D
        active = active & ((C_R + C_D) < PERCEIVED_VALUE)

    # Calculate dynamic stats dynamically without reading DataFrames
    batch_sum_v = cum_v_res.sum().item()
    batch_sum_v_sq = (cum_v_res ** 2).sum().item()
    
    step_cpu = step.cpu().numpy()
    survived_mask = step_cpu == MAX_STEPS
    revolution_mask = (~survived_mask) & (last_contrib_R.cpu().numpy() > last_contrib_D.cpu().numpy())
    
    batch_surv = survived_mask.sum()
    batch_rev = revolution_mask.sum()
    batch_crack = n_games - batch_surv - batch_rev

    term_status = np.full(n_games, "Crackdown", dtype=object)
    term_status[survived_mask] = "Survived"
    term_status[revolution_mask] = "Revolution"

    df_batch = pd.DataFrame({
        'Sim_Type': np.int8(sim_type),
        'Case': pd.Categorical([case if case else 'Random'] * n_games),
        'Archetype': pd.Categorical([archetype if archetype else 'None'] * n_games),
        'Init_P_R': p_R_base.cpu().numpy().astype(np.float32),
        'Init_P_D': p_D_base.cpu().numpy().astype(np.float32),
        'Steps': step_cpu.astype(np.int32),
        'Cum_V_res': cum_v_res.cpu().numpy().astype(np.float32),
        'Terminal_Status': pd.Categorical(term_status),
        'Avg_P_R': (sum_p_R / step).cpu().numpy().astype(np.float32),
        'Avg_P_D': (sum_p_D / step).cpu().numpy().astype(np.float32),
        'Avg_C_R': (sum_c_R / step).cpu().numpy().astype(np.float32),
        'Avg_C_D': (sum_c_D / step).cpu().numpy().astype(np.float32),
        'Avg_Pen_R': (sum_pen_R / step).cpu().numpy().astype(np.float32),
        'Avg_Pen_D': (sum_pen_D / step).cpu().numpy().astype(np.float32)
    })
    
    stats = {
        'count': n_games,
        'sum_v': batch_sum_v,
        'sum_v_sq': batch_sum_v_sq,
        'surv': int(batch_surv),
        'rev': int(batch_rev),
        'crack': int(batch_crack)
    }

    if queue is not None:
        queue.put(df_batch)

    return stats

def run_game(sim_type, case=None, archetype=None):
    random.seed()
    C_R, C_D = 0.0, 0.0
    history = {}
    cum_v_res = 0.0
    
    prob_r_list, prob_d_list = [], []
    cost_r_list, cost_d_list = [], []
    pen_r_list, pen_d_list = [], []
    
    if case == 'DD':
        case_p_R, case_p_D = random.uniform(0, 0.5), random.uniform(0, 0.5)
    elif case == 'DH':
        case_p_R, case_p_D = random.uniform(0, 0.5), random.uniform(0.5, 1.0)
    elif case == 'HD':
        case_p_R, case_p_D = random.uniform(0.5, 1.0), random.uniform(0, 0.5)
    elif case == 'HH':
        case_p_R, case_p_D = random.uniform(0.5, 1.0), random.uniform(0.5, 1.0)
    else: # 'RR' or 'None' (Random)
        case_p_R, case_p_D = random.uniform(0, 1.0), random.uniform(0, 1.0)

    p_R_base, p_D_base = case_p_R, case_p_D
    
    if sim_type in [3, 4]:
        p_R_base, p_D_base = random.uniform(0.0, 1.0), random.uniform(0.0, 1.0)
    elif sim_type in [5, 6]:
        p_R_base, p_D_base = random.uniform(0.50001, 1.0), random.uniform(0.0, 1.0)
    elif sim_type in [7, 8]:
        p_R_base, p_D_base = random.uniform(0.0, 1.0), random.uniform(0.50001, 1.0)

    p_R_hawk_if_H = random.uniform(0.50001, 1.0)  
    p_R_hawk_if_D = random.uniform(0.0, 0.49999)  
    p_D_hawk_if_H = random.uniform(0.50001, 1.0)  
    p_D_hawk_if_D = random.uniform(0.0, 0.49999)  

    step = 1
    dc_R = random.uniform(0, MAX_COST_PER_STEP)
    dc_D = random.uniform(0, MAX_COST_PER_STEP)
    
    contrib_R, contrib_D = dc_R, -dc_D
    
    cost_r_list.append(contrib_R / 2.0)
    cost_d_list.append(contrib_D / 2.0)
    prob_r_list.append(1.0)
    prob_d_list.append(0.0)
    pen_r_list.append(1.0)
    pen_d_list.append(1.0)
    
    prev_act_R, prev_act_D = 'H', 'D'
    
    delta_C = (contrib_R + contrib_D) / 2.0
    C_R, C_D = max(0.0, C_R + delta_C), max(0.0, C_D + delta_C)
    
    v_res = PERCEIVED_VALUE - (C_R + C_D)
    cum_v_res += max(0.0, v_res)
    
    if sim_type >= 3 and sim_type != 9:
        next_penalty = max(0.0, v_res) / PERCEIVED_VALUE
        if archetype == 'Equals':
            next_pen_R, next_pen_D = next_penalty, next_penalty
        elif archetype == 'Cowards':
            next_pen_R, next_pen_D = next_penalty ** 2, next_penalty ** 2
        elif archetype == 'Fools':
            next_pen_R, next_pen_D = next_penalty ** 0.5, next_penalty ** 0.5
        elif archetype == 'Brinksmen':
            next_pen_R, next_pen_D = next_penalty, 1.0 - (1.0 - next_penalty) / 3.0
        elif archetype == 'Tyrants':
            next_pen_R, next_pen_D = 1.0 - (1.0 - next_penalty) / 3.0, next_penalty
        else:
            next_pen_R, next_pen_D = 1.0, 1.0
    else:
        next_pen_R, next_pen_D = 1.0, 1.0

    while (C_R + C_D) < PERCEIVED_VALUE and step < MAX_STEPS:
        step += 1
        v_res = PERCEIVED_VALUE - (C_R + C_D)
        
        if sim_type >= 3 and sim_type != 9:
            penalty = v_res / PERCEIVED_VALUE
            if archetype == 'Equals':
                mul_R, mul_D = penalty, penalty
            elif archetype == 'Cowards':
                mul_R, mul_D = penalty ** 2, penalty ** 2
            elif archetype == 'Fools':
                mul_R, mul_D = penalty ** 0.5, penalty ** 0.5
            elif archetype == 'Brinksmen':
                mul_R, mul_D = penalty, 1.0 - (1.0 - penalty)/3.0
            elif archetype == 'Tyrants':
                mul_R, mul_D = 1.0 - (1.0 - penalty)/3.0, penalty
            else:
                mul_R, mul_D = 1.0, 1.0
        else:
            mul_R, mul_D = 1.0, 1.0

        if sim_type == 1:
            p_R_curr = random.uniform(0, 1.0)
            p_D_curr = random.uniform(0, 1.0)
        elif sim_type == 2:
            p_R_curr = p_R_base
            p_D_curr = p_D_base
        else:
            if step == 2:
                base_R, base_D = p_R_base, p_D_base
            else:
                if sim_type in [3, 5, 7]:
                    base_R, base_D = p_R_base, p_D_base
                elif sim_type in [4, 9]:
                    base_R = p_R_hawk_if_H if prev_act_D == 'H' else p_R_hawk_if_D
                    base_D = p_D_hawk_if_H if prev_act_R == 'H' else p_D_hawk_if_D
                elif sim_type == 6:
                    base_R = p_R_base
                    base_D = p_D_hawk_if_H if prev_act_R == 'H' else p_D_hawk_if_D
                elif sim_type == 8:
                    base_R = p_R_hawk_if_H if prev_act_D == 'H' else p_R_hawk_if_D
                    base_D = p_D_base
            
            p_R_curr = max(0.01, base_R * mul_R)
            p_D_curr = max(0.01, base_D * mul_D)

        prob_r_list.append(p_R_curr)
        prob_d_list.append(p_D_curr)
        pen_r_list.append(mul_R)
        pen_d_list.append(mul_D)

        act_R = 'H' if random.random() < p_R_curr else 'D'
        act_D = 'H' if random.random() < p_D_curr else 'D'
        
        prev_act_R, prev_act_D = act_R, act_D
        
        dc_R = random.uniform(0, MAX_COST_PER_STEP)
        dc_D = random.uniform(0, MAX_COST_PER_STEP)
        
        contrib_R = dc_R if act_R == 'H' else -dc_R
        contrib_D = dc_D if act_D == 'H' else -dc_D
        
        cost_r_list.append(contrib_R / 2.0)
        cost_d_list.append(contrib_D / 2.0)
        
        delta_C = (contrib_R + contrib_D) / 2.0
        
        C_R = max(0.0, C_R + delta_C)
        C_D = max(0.0, C_D + delta_C)
        
        v_res = max(0.0, PERCEIVED_VALUE - (C_R + C_D))
        cum_v_res += v_res
        
    if step == MAX_STEPS:
        term_status = "Survived"
    else:
        if_r_lost = contrib_R > contrib_D
        if if_r_lost:
            term_status = "Revolution"
        else:
            term_status = "Crackdown"
            
    result = {
        'Sim_Type': sim_type,
        'Case': case if case else 'Random',
        'Archetype': archetype if archetype else 'None',
        'Init_P_R': p_R_base,
        'Init_P_D': p_D_base,
        'Steps': step,
        'Cum_V_res': cum_v_res,
        'Terminal_Status': term_status,
        'Avg_P_R': np.mean(prob_r_list),
        'Avg_P_D': np.mean(prob_d_list),
        'Avg_C_R': np.mean(cost_r_list),
        'Avg_C_D': np.mean(cost_d_list),
        'Avg_Pen_R': np.mean(pen_r_list),
        'Avg_Pen_D': np.mean(pen_d_list)
    }
    return result, history

def _run_game_task_wrapper(args):
    sim_type, case, archetype = args
    return run_game(sim_type, case, archetype)

_worker_pool = None

def init_persistent_workers():
    global _worker_pool
    if _worker_pool is None:
        _worker_pool = multiprocessing.Pool(processes=max(1, multiprocessing.cpu_count() - 1))

def shutdown_persistent_workers():
    global _worker_pool
    if _worker_pool is not None:
        _worker_pool.close()
        _worker_pool.join()
        _worker_pool = None

def format_time(seconds):
    m, s = divmod(int(seconds), 60)
    h, m = divmod(m, 60)
    if h > 0: return f"{h:02d}:{m:02d}:{s:02d}"
    return f"{m:02d}:{s:02d}"

def progress_worker_process(q, total, desc="Executing"):
    import time
    import queue as q_module
    start_time = time.time()
    completed = 0
    speed = 0
    eta = 0
    eta_lapsed = 0.0
    eta_lapsed_start = 0.0

    while True:
        elapsed = time.time() - start_time

        # Look for updates
        try:
            msg = q.get(timeout=1.0)
            if msg == 'DONE':
                completed = total
                break
            elif isinstance(msg, int):
                completed = msg
                speed = completed / elapsed if elapsed else 0 # Only update when new data is returned
                eta = (total - completed) / speed if speed else 0
                eta_lapsed_start = time.time()
        except q_module.Empty:
            pass

        # Dynamically update Eta
        eta_lapsed =  time.time() - eta_lapsed_start if eta_lapsed_start else 0.0
        current_eta = max(0.0, eta - eta_lapsed)
        
        bar_length = 20  # Reduced to prevent terminal wrapping
        filled = int(bar_length * completed / total) if total > 0 else 0
        bar = '#' * filled + '-' * (bar_length - filled)
        percent = (completed / total) * 100 if total > 0 else 0

        # Shortened labels and added trailing spaces to overwrite artifacts cleanly
        print(f"\r{desc}: [{bar}] {completed}/{total} ({percent:.1f}%) | {speed:.0f} sim/s | Ela: {format_time(elapsed)} | ETA: {format_time(current_eta)}   ", end="", flush=True)
        
    elapsed = time.time() - start_time
    speed = total / elapsed if elapsed > 0 else 0
    bar = '#' * 20
    print(f"\r{desc}: [{bar}] {total}/{total} (100.0%) | {speed:.0f} sim/s | Ela: {format_time(elapsed)} | ETA: 00:00   \n", end="", flush=True)

def run_parallel_jobs(jobs, queue, agg_stats, desc="Executing"):
    args_list = [j[1] for j in jobs]
    if _worker_pool:
        total = len(args_list)
        
        prog_queue = multiprocessing.Manager().Queue()
        prog_proc = multiprocessing.Process(target=progress_worker_process, args=(prog_queue, total, desc))
        prog_proc.start()
        
        batch = []
        last_put_time = time.time()
        
        for i, res in enumerate(_worker_pool.imap_unordered(_run_game_task_wrapper, args_list), 1):
            result_dict, hist = res
            batch.append(result_dict)
            
            # Aggregate stats without memory load
            key = (result_dict['Sim_Type'], result_dict['Case'], result_dict['Archetype'])
            if key not in agg_stats:
                 agg_stats[key] = {'count':0, 'sum_v':0.0, 'sum_v_sq':0.0, 'surv':0, 'rev':0, 'crack':0}
            
            agg_stats[key]['count'] += 1
            agg_stats[key]['sum_v'] += result_dict['Cum_V_res']
            agg_stats[key]['sum_v_sq'] += result_dict['Cum_V_res']**2
            
            term = result_dict['Terminal_Status']
            if term == 'Survived': agg_stats[key]['surv'] += 1
            elif term == 'Revolution': agg_stats[key]['rev'] += 1
            elif term == 'Crackdown': agg_stats[key]['crack'] += 1
            
            if len(batch) >= 10000:
                if queue: queue.put(pd.DataFrame(batch))
                batch = []
            
            curr_time = time.time()
            if curr_time - last_put_time >= 0.5 or i == total:
                prog_queue.put(i)
                last_put_time = curr_time
  
        if batch and queue:
            queue.put(pd.DataFrame(batch))
            
        prog_queue.put('DONE')
        prog_proc.join()
        return

def run_all_simulations():
    manager = multiprocessing.Manager()
    queue = manager.Queue()

    summary_file = os.path.join(TEMP_DIR, FILE_SUMMARY)
    feat_file = os.path.join(TEMP_DIR, FILE_FEATURES)

    summary_cols = ['Sim_Type', 'Case', 'Archetype', 'Init_P_R', 'Init_P_D', 'Steps', 'Cum_V_res', 'Terminal_Status', 'Avg_P_R', 'Avg_P_D', 'Avg_C_R', 'Avg_C_D', 'Avg_Pen_R', 'Avg_Pen_D']
    feat_cols = ['Sim_Type', 'Case', 'Archetype', 'Terminal_Status', 'Avg_P_R', 'Avg_P_D', 'Avg_C_R', 'Avg_C_D', 'Avg_Pen_R', 'Avg_Pen_D', 'Cum_V_res']
    pd.DataFrame(columns=summary_cols).to_csv(summary_file, index=False)
    pd.DataFrame(columns=feat_cols).to_csv(feat_file, index=False)

    writer_proc = multiprocessing.Process(target=csv_writer_process, args=(queue, summary_file, feat_file))
    writer_proc.start()

    agg_stats = {}

    if HAS_CUDA:
        print(f"CUDA GPU hardware detected: [{torch.cuda.get_device_name(0)}]")
        print(f"Running vectorized GPU Monte Carlo simulations (Scaling N_REPS = {N_REPS})...")
        
        cases = ['DD', 'DH', 'HD', 'HH', 'RR']
        archetypes = ['Equals', 'Cowards', 'Fools', 'Brinksmen', 'Tyrants']
        
        # Calculate strict total jobs for progress bar
        total_sims = N_REPS * 10 + len(cases) * N_REPS * 2 + 6 * len(archetypes) * N_REPS * 5
        jobs_completed = 0
        
        prog_queue = manager.Queue()
        prog_proc = multiprocessing.Process(target=progress_worker_process, args=(prog_queue, total_sims, "Executing GPU Pipeline"))
        prog_proc.start()

        def execute_gpu_batch(sim_type, case, arch, total_n):
            nonlocal jobs_completed, agg_stats
            chunks = []
            MAX_GPU_BATCH = 1000000  # Avoid tensor exhaustion
            rem = total_n
            while rem > 0:
                chunks.append(min(rem, MAX_GPU_BATCH))
                rem -= chunks[-1]

            for n in chunks:
                stats = run_games_cuda(sim_type, case, arch, n, queue)
                
                key = (sim_type, case, arch)
                if key not in agg_stats:
                    agg_stats[key] = {'count':0, 'sum_v':0.0, 'sum_v_sq':0.0, 'surv':0, 'rev':0, 'crack':0}
                agg_stats[key]['count'] += stats['count']
                agg_stats[key]['sum_v'] += stats['sum_v']
                agg_stats[key]['sum_v_sq'] += stats['sum_v_sq']
                agg_stats[key]['surv'] += stats['surv']
                agg_stats[key]['rev'] += stats['rev']
                agg_stats[key]['crack'] += stats['crack']

                jobs_completed += n
                prog_queue.put(jobs_completed)

        execute_gpu_batch(1, None, None, N_REPS * 10)
        for case in cases:
            execute_gpu_batch(2, case, None, N_REPS)
            execute_gpu_batch(9, case, None, N_REPS)
        for sim_type in range(3, 9):
            for arch in archetypes:
                execute_gpu_batch(sim_type, None, arch, N_REPS * 5)
        
        prog_queue.put('DONE')
        prog_proc.join()
    else:
        print("CUDA GPU not detected. Falling back to persistent CPU worker multiprocessing...")
        cases = ['DD', 'DH', 'HD', 'HH', 'RR']
        archetypes = ['Equals', 'Cowards', 'Fools', 'Brinksmen', 'Tyrants']
        jobs = []

        for _ in range(N_REPS * 10): jobs.append(('SIMULATE', (1, None, None)))
        for case in cases:
            for _ in range(N_REPS):
                jobs.append(('SIMULATE', (2, case, None)))
                jobs.append(('SIMULATE', (9, case, None)))
        for sim_type in range(3, 9):
            for arch in archetypes:
                for _ in range(N_REPS * 5):
                    jobs.append(('SIMULATE', (sim_type, None, arch)))

        run_parallel_jobs(jobs, queue, agg_stats, desc="Executing CPU Simulations")

    print("\nFinalizing Background Write Stream...")
    queue.put(None)
    writer_proc.join()
    
    # Construct an aggregated DataFrame strictly for visualizations, totally sidestepping RAM constraints
    records = []
    for (sim, case, arch), s in agg_stats.items():
        count = s['count']
        if count == 0: continue
        mean = s['sum_v'] / count
        var = (s['sum_v_sq'] / count) - (mean ** 2)
        std = np.sqrt(max(0, var))
        records.append({
            'Sim_Type': sim,
            'Case': case if case else 'Random',
            'Archetype': arch if arch else 'None',
            'Mean_V_res': mean,
            'Std_V_res': std,
            'Total_Count': count,
            'Crackdown_Count': s['crack'],
            'Revolution_Count': s['rev'],
            'Survived_Count': s['surv']
        })
    df_agg = pd.DataFrame(records)
    df_agg.loc[df_agg['Sim_Type'].isin([2, 9]), 'Archetype'] = 'Control'

    return df_agg, pd.DataFrame()

# ==========================================
# 4. VISUALIZATIONS
# ==========================================
def compute_paired_significance(sub_base, sub_t4t, metric_col):
    if sub_base.empty or sub_t4t.empty:
        return 1.0, "ns"
        
    if metric_col == 'Is_Terminated':
        term_b = sub_base['Crackdown_Count'].values[0] + sub_base['Revolution_Count'].values[0]
        surv_b = sub_base['Survived_Count'].values[0]
        term_t = sub_t4t['Crackdown_Count'].values[0] + sub_t4t['Revolution_Count'].values[0]
        surv_t = sub_t4t['Survived_Count'].values[0]
        contingency = [[term_b, surv_b], [term_t, surv_t]]

        try:
            res = stats.chi2_contingency(contingency)
            p_val = res.pvalue
        except Exception:
            _, p_val = stats.fisher_exact(contingency)
    else:
        mb = sub_base['Mean_V_res'].values[0]
        sb = sub_base['Std_V_res'].values[0]
        nb = sub_base['Total_Count'].values[0]
        mt = sub_t4t['Mean_V_res'].values[0]
        st = sub_t4t['Std_V_res'].values[0]
        nt = sub_t4t['Total_Count'].values[0]
        if nb < 2 or nt < 2: return 1.0, "ns"
        _, p_val = stats.ttest_ind_from_stats(mb, sb, nb, mt, st, nt, equal_var=False)

    if np.isnan(p_val): p_val = 1.0
    if p_val < 0.001: return p_val, "***"
    elif p_val < 0.01: return p_val, "**"
    elif p_val < 0.05: return p_val, "*"
    else: return p_val, "ns"

def plot_sim_pair_dual_canvas(df_archetypes, sim_base, sim_t4t, pair_label, output_folder):
    if sim_base == 2 and sim_t4t == 9:
        x_categories = ['DD', 'DH', 'HD', 'HH', 'RR']
        group_col = 'Case'
        x_label = 'Starting Case Quadrant'
    else:
        std_order = STANDARD_ARCHETYPE_ORDER
        detected = sorted(df_archetypes['Archetype'].unique().tolist())
        x_categories = [a for a in std_order if a in detected and a != 'Control'] + [a for a in detected if a not in std_order and a != 'Control']
        if 'None' in x_categories:
            x_categories.remove('None')
        group_col = 'Archetype'
        x_label = 'National Archetype'

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 6))

    n_cats = len(x_categories)
    bar_width = 0.35
    x_indices = np.arange(n_cats)

    color_base_crack, color_base_rev = '#1f77b4', '#6baed6'
    color_t4t_crack, color_t4t_rev = '#d95f02', '#fdbf6f'

    means_base_p, stds_base_p, means_t4t_p, stds_t4t_p, sig_p = [], [], [], [], []
    crack_base_t, rev_base_t, total_base_t = [], [], []
    crack_t4t_t, rev_t4t_t, total_t4t_t = [], [], []
    cnt_crack_b, cnt_rev_b, cnt_tot_b, cnt_crack_t, cnt_rev_t, cnt_tot_t, sig_t = [], [], [], [], [], [], []

    for cat in x_categories:
        sub_base = df_archetypes[(df_archetypes[group_col] == cat) & (df_archetypes['Sim_Type'] == sim_base)]
        sub_t4t = df_archetypes[(df_archetypes[group_col] == cat) & (df_archetypes['Sim_Type'] == sim_t4t)]

        if not sub_base.empty:
            means_base_p.append(sub_base['Mean_V_res'].values[0])
            stds_base_p.append(sub_base['Std_V_res'].values[0])
            cb_c = sub_base['Crackdown_Count'].values[0]
            cb_r = sub_base['Revolution_Count'].values[0]
            n_tot_b = sub_base['Total_Count'].values[0]
        else:
            means_base_p.append(0.0); stds_base_p.append(0.0); cb_c = cb_r = n_tot_b = 0

        if not sub_t4t.empty:
            means_t4t_p.append(sub_t4t['Mean_V_res'].values[0])
            stds_t4t_p.append(sub_t4t['Std_V_res'].values[0])
            ct_c = sub_t4t['Crackdown_Count'].values[0]
            ct_r = sub_t4t['Revolution_Count'].values[0]
            n_tot_t = sub_t4t['Total_Count'].values[0]
        else:
            means_t4t_p.append(0.0); stds_t4t_p.append(0.0); ct_c = ct_r = n_tot_t = 0
            
        _, stars_p = compute_paired_significance(sub_base, sub_t4t, 'Cum_V_res')
        sig_p.append(stars_p)

        p_cb_c, p_cb_r = (cb_c / n_tot_b if n_tot_b > 0 else 0.0), (cb_r / n_tot_b if n_tot_b > 0 else 0.0)
        p_ct_c, p_ct_r = (ct_c / n_tot_t if n_tot_t > 0 else 0.0), (ct_r / n_tot_t if n_tot_t > 0 else 0.0)

        crack_base_t.append(p_cb_c); rev_base_t.append(p_cb_r); total_base_t.append(p_cb_c + p_cb_r)
        cnt_crack_b.append(cb_c); cnt_rev_b.append(cb_r); cnt_tot_b.append(n_tot_b)
        crack_t4t_t.append(p_ct_c); rev_t4t_t.append(p_ct_r); total_t4t_t.append(p_ct_c + p_ct_r)
        cnt_crack_t.append(ct_c); cnt_rev_t.append(ct_r); cnt_tot_t.append(n_tot_t)
        
        _, stars_t = compute_paired_significance(sub_base, sub_t4t, 'Is_Terminated')
        sig_t.append(stars_t)

    err_base_p = [[min(m, s) for m, s in zip(means_base_p, stds_base_p)], stds_base_p]
    err_t4t_p = [[min(m, s) for m, s in zip(means_t4t_p, stds_t4t_p)], stds_t4t_p]

    rects1_p = ax1.bar(x_indices - bar_width/2, means_base_p, bar_width, yerr=err_base_p, capsize=4,
                       color='#2b5c8f', edgecolor='black', linewidth=0.7, label=SIM_LABELS[sim_base])
    rects2_p = ax1.bar(x_indices + bar_width/2, means_t4t_p, bar_width, yerr=err_t4t_p, capsize=4,
                       color='#e05d44', edgecolor='black', linewidth=0.7, label=SIM_LABELS[sim_t4t])

    for rect, val in zip(rects1_p, means_base_p):
        if val > 0:
            ax1.text(rect.get_x() + rect.get_width()/2., val / 2., f"{val:.0f}",
                     ha='center', va='center', color='white', fontweight='bold', fontsize=8.5, rotation=90)
    for rect, val in zip(rects2_p, means_t4t_p):
        if val > 0:
            ax1.text(rect.get_x() + rect.get_width()/2., val / 2., f"{val:.0f}",
                     ha='center', va='center', color='white', fontweight='bold', fontsize=8.5, rotation=90)

    for x_i, mb, mt, sb, st, stars in zip(x_indices, means_base_p, means_t4t_p, stds_base_p, stds_t4t_p, sig_p):
        max_height = max(mb + sb, mt + st)
        y_pos = max_height + max_height * 0.03 + 20
        font_color, font_weight = ('#d9534f', 'bold') if stars != 'ns' else ('#666666', 'normal')
        ax1.text(x_i, y_pos, stars, ha='center', va='bottom', fontsize=11, color=font_color, fontweight=font_weight)

    ax1.set_ylim(0, 8000)
    ax1.set_title('(A) Prosperity Scores (Mean Cum_V_res)', fontweight='bold', fontsize=11, pad=10)
    ax1.set_xticks(x_indices)
    ax1.set_xticklabels(x_categories, fontweight='bold', fontsize=9.5)
    ax1.set_xlabel(x_label, fontweight='bold')
    ax1.set_ylabel('Mean Cum_V_res', fontweight='bold', fontsize=10)
    ax1.grid(axis='y', linestyle='--', alpha=0.5)
    ax1.legend(loc='upper right', frameon=True, facecolor='white', framealpha=0.9, fontsize=8.0)

    x_base, x_t4t = x_indices - bar_width/2, x_indices + bar_width/2
    ax2.bar(x_base, crack_base_t, bar_width, label=f"{SIM_LABELS[sim_base]} (Crackdown)", color=color_base_crack, edgecolor='black', linewidth=0.7)
    ax2.bar(x_base, rev_base_t, bar_width, bottom=crack_base_t, label=f"{SIM_LABELS[sim_base]} (Revolution)", color=color_base_rev, edgecolor='black', linewidth=0.7)
    ax2.bar(x_t4t, crack_t4t_t, bar_width, label=f"{SIM_LABELS[sim_t4t]} (Crackdown)", color=color_t4t_crack, edgecolor='black', linewidth=0.7)
    ax2.bar(x_t4t, rev_t4t_t, bar_width, bottom=crack_t4t_t, label=f"{SIM_LABELS[sim_t4t]} (Revolution)", color=color_t4t_rev, edgecolor='black', linewidth=0.7)

    for i in range(n_cats):
        # Base Bar Labels
        pc_b, nc_b = crack_base_t[i], cnt_crack_b[i]
        pr_b, nr_b = rev_base_t[i], cnt_rev_b[i]
        pt_b, nt_b = total_base_t[i], cnt_tot_b[i]

        if pc_b >= 0.04:
            ax2.text(x_base[i], pc_b / 2.0, f"{pc_b*100:.0f}%\n({nc_b})",
                     ha='center', va='center', color='white', fontweight='bold', fontsize=6.5)
        if pr_b >= 0.04:
            ax2.text(x_base[i], pc_b + pr_b / 2.0, f"{pr_b*100:.0f}%\n({nr_b})",
                     ha='center', va='center', color='#111111', fontweight='bold', fontsize=6.5)
        if pt_b > 0:
            ax2.text(x_base[i], pt_b + 0.01, f"{pt_b*100:.1f}%\n(n={nt_b})",
                     ha='center', va='bottom', color='black', fontweight='bold', fontsize=7.5)

        # T4T Bar Labels
        pc_t, nc_t = crack_t4t_t[i], cnt_crack_t[i]
        pr_t, nr_t = rev_t4t_t[i], cnt_rev_t[i]
        pt_t, nt_t = total_t4t_t[i], cnt_tot_t[i]

        if pc_t >= 0.04:
            ax2.text(x_t4t[i], pc_t / 2.0, f"{pc_t*100:.0f}%\n({nc_t})",
                     ha='center', va='center', color='white', fontweight='bold', fontsize=6.5)
        if pr_t >= 0.04:
            ax2.text(x_t4t[i], pc_t + pr_t / 2.0, f"{pr_t*100:.0f}%\n({nr_t})",
                     ha='center', va='center', color='#111111', fontweight='bold', fontsize=6.5)
        if pt_t > 0:
            ax2.text(x_t4t[i], pt_t + 0.01, f"{pt_t*100:.1f}%\n(n={nt_t})",
                     ha='center', va='bottom', color='black', fontweight='bold', fontsize=7.5)

        max_h = max(total_base_t[i], total_t4t_t[i])
        stars = sig_t[i]
        if max_h > 0:
            y_pos = max_h + 0.07
            font_color, font_weight = ('#d9534f', 'bold') if stars != 'ns' else ('#666666', 'normal')
            ax2.text(x_indices[i], y_pos, stars, ha='center', va='bottom', fontsize=11, color=font_color, fontweight=font_weight)

    max_bar_val = max(max(total_base_t), max(total_t4t_t)) if total_base_t and total_t4t_t else 0.5
    ax2.set_ylim(0, max(max_bar_val * 1.35, 0.15))
    ax2.set_title('(B) Termination Proportions (Crackdown + Revolution)', fontweight='bold', fontsize=11, pad=10)
    ax2.set_xticks(x_indices)
    ax2.set_xticklabels(x_categories, fontweight='bold', fontsize=9.5)
    ax2.set_xlabel(x_label, fontweight='bold')
    ax2.set_ylabel('Termination Proportion (0.0 - 1.0)', fontweight='bold', fontsize=10)
    ax2.grid(axis='y', linestyle='--', alpha=0.5)
    ax2.legend(loc='upper left', frameon=True, facecolor='white', framealpha=0.9, fontsize=8.0)

    plt.suptitle(f'Paired Analysis: {pair_label}', fontweight='bold', fontsize=13, y=1.05)
    plt.tight_layout()

    file_path = os.path.join(output_folder, f'dual_canvas_sim{sim_base}_vs_sim{sim_t4t}_archetypes.png')
    plt.savefig(file_path, dpi=300, bbox_inches='tight')
    plt.close()
    return file_path

def plot_cross_archetype_dual_canvas(df_archetypes, output_folder):
    std_order = STANDARD_ARCHETYPE_ORDER
    detected = sorted(df_archetypes['Archetype'].unique().tolist())
    detected_archetypes = [a for a in std_order if a in detected and a != 'Control'] + [a for a in detected if a not in std_order and a != 'Control']
    
    if 'None' in detected_archetypes:
        detected_archetypes.remove('None')

    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(16, 12))

    n_archs = len(detected_archetypes)
    ordered_sims = [2, 9, 3, 4, 5, 6, 7, 8]
    n_sims = len(ordered_sims)
    bar_width = 0.10
    x_indices = np.arange(n_archs)
    
    cluster_width = n_sims * bar_width + 3 * 0.02
    start_x = -cluster_width / 2 + bar_width / 2

    sim_colors_crack = {
        2: '#7570b3', 9: '#9e9ac8', 3: '#1f77b4', 4: '#3182bd',
        5: '#d95f02', 6: '#e6550d', 7: '#2ca02c', 8: '#31a354'
    }
    sim_colors_rev = {
        2: '#bcbddc', 9: '#dadaeb', 3: '#9ecae1', 4: '#c6dbef',
        5: '#fdbf6f', 6: '#fdd0a2', 7: '#a1d99b', 8: '#e5f5e0'
    }

    all_max_heights = []

    for i, sim_type in enumerate(ordered_sims):
        pair_idx = i // 2
        offset = start_x + i * bar_width + pair_idx * 0.02
        
        means_p, stds_p = [], []
        crack_means, rev_means, tot_means = [], [], []

        for arch in detected_archetypes:
            if sim_type in [2, 9]:
                # Combine all cases for Sims 2 and 9 explicitly as baseline global clusters
                sub = df_archetypes[df_archetypes['Sim_Type'] == sim_type]
            else:
                sub = df_archetypes[(df_archetypes['Archetype'] == arch) & (df_archetypes['Sim_Type'] == sim_type)]
            
            if not sub.empty:
                cnt_tot_n = sub['Total_Count'].sum()
                means_p.append(sub['Mean_V_res'].mean())  # mean of means across cases
                stds_p.append(sub['Std_V_res'].mean())
                p_c = sub['Crackdown_Count'].sum() / cnt_tot_n if cnt_tot_n > 0 else 0.0
                p_r = sub['Revolution_Count'].sum() / cnt_tot_n if cnt_tot_n > 0 else 0.0
            else:
                means_p.append(0.0); stds_p.append(0.0)
                p_c = p_r = 0.0
                
            crack_means.append(p_c)
            rev_means.append(p_r)
            tot_means.append(p_c + p_r)
            all_max_heights.append(p_c + p_r)

        err_p = [[min(m, s) for m, s in zip(means_p, stds_p)], stds_p]

        ax1.bar(x_indices + offset, means_p, bar_width, yerr=err_p, capsize=3,
                color=sim_colors_crack[sim_type], edgecolor='black', linewidth=0.6, label=SIM_LABELS[sim_type])
                
        ax2.bar(x_indices + offset, crack_means, bar_width, label=f"{SIM_LABELS[sim_type]} (Crackdown)",
                color=sim_colors_crack[sim_type], edgecolor='black', linewidth=0.6)
        ax2.bar(x_indices + offset, rev_means, bar_width, bottom=crack_means, label=f"{SIM_LABELS[sim_type]} (Revolution)",
                color=sim_colors_rev[sim_type], edgecolor='black', linewidth=0.6)

    ax1.set_ylabel('Mean Cumulative Value (Cum_V_res)', fontweight='bold')
    ax1.set_xlabel('National Archetype', fontweight='bold')
    ax1.set_title('(A) Prosperity Scores across Archetypes (All Sims)', fontweight='bold', pad=10)
    ax1.set_xticks(x_indices)
    ax1.set_xticklabels(detected_archetypes, fontweight='bold', fontsize=10)
    ax1.grid(axis='y', linestyle='--', alpha=0.5)
    ax1.legend(title='Simulation Framework', loc='center left', bbox_to_anchor=(1.01, 0.5), frameon=True, facecolor='white', framealpha=0.9, fontsize=7.5)

    max_h = max(all_max_heights) if all_max_heights else 0.5
    ax2.set_ylim(0, max(max_h * 1.35, 0.15))
    ax2.set_ylabel('Termination Proportion (Crackdown + Revolution)', fontweight='bold')
    ax2.set_xlabel('National Archetype', fontweight='bold')
    ax2.set_title('(B) Stacked Termination Proportions across Archetypes (All Sims)', fontweight='bold', pad=10)
    ax2.set_xticks(x_indices)
    ax2.set_xticklabels(detected_archetypes, fontweight='bold', fontsize=10)
    ax2.grid(axis='y', linestyle='--', alpha=0.5)
    ax2.legend(title='Simulation Framework & Outcome', loc='center left', bbox_to_anchor=(1.01, 0.5), fontsize=7.5, framealpha=0.9)

    plt.suptitle('Consolidated Overview: Prosperity vs. Societal Stability across Archetypes', fontweight='bold', fontsize=14, y=1.05)
    plt.tight_layout(rect=[0, 0, 0.85, 1])

    file_path = os.path.join(output_folder, 'dual_canvas_all_sims_cross_archetype.png')
    plt.savefig(file_path, dpi=300, bbox_inches='tight')
    plt.close()
    return file_path

def plot_individual_archetype_dual_canvas(df_archetypes, archetype_name, output_folder):
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 6))

    pair_labels = ['Control\n(Sim 2 vs 9)', 'Neutral\n(Sim 3 vs 4)', 'Agg Ruler\n(Sim 5 vs 6)', 'Agg Ruled\n(Sim 7 vs 8)']
    n_pairs = len(SIM_PAIRS)
    bar_width = 0.35
    x_indices = np.arange(n_pairs)

    color_base_crack, color_base_rev = '#1f77b4', '#6baed6'
    color_t4t_crack, color_t4t_rev = '#d95f02', '#fdbf6f'

    means_base_p, stds_base_p, means_t4t_p, stds_t4t_p, sig_p = [], [], [], [], []
    crack_base_t, rev_base_t, total_base_t = [], [], []
    crack_t4t_t, rev_t4t_t, total_t4t_t = [], [], []
    cnt_crack_b, cnt_rev_b, cnt_tot_b = [], [], []
    cnt_crack_t, cnt_rev_t, cnt_tot_t = [], [], []
    sig_t = []

    for sim_base, sim_t4t, _ in SIM_PAIRS:
        # Sums/means over potential cases
        if sim_base == 2 and sim_t4t == 9:
            sub_base = df_archetypes[df_archetypes['Sim_Type'] == 2]
            sub_t4t = df_archetypes[df_archetypes['Sim_Type'] == 9]
        else:
            sub_base = df_archetypes[(df_archetypes['Archetype'] == archetype_name) & (df_archetypes['Sim_Type'] == sim_base)]
            sub_t4t = df_archetypes[(df_archetypes['Archetype'] == archetype_name) & (df_archetypes['Sim_Type'] == sim_t4t)]

        # Consolidate baseline case groupings
        if not sub_base.empty:
            mb = sub_base['Mean_V_res'].mean()
            sb = sub_base['Std_V_res'].mean()
            means_base_p.append(mb)
            stds_base_p.append(sb)
            n_tot_b = sub_base['Total_Count'].sum()
            cb_c = sub_base['Crackdown_Count'].sum()
            cb_r = sub_base['Revolution_Count'].sum()
        else:
            means_base_p.append(0.0); stds_base_p.append(0.0); cb_c = cb_r = n_tot_b = 0
            
        if not sub_t4t.empty:
            mt = sub_t4t['Mean_V_res'].mean()
            st = sub_t4t['Std_V_res'].mean()
            means_t4t_p.append(mt)
            stds_t4t_p.append(st)
            n_tot_t = sub_t4t['Total_Count'].sum()
            ct_c = sub_t4t['Crackdown_Count'].sum()
            ct_r = sub_t4t['Revolution_Count'].sum()
        else:
            means_t4t_p.append(0.0); stds_t4t_p.append(0.0); ct_c = ct_r = n_tot_t = 0
            
        _, stars_p = compute_paired_significance(sub_base, sub_t4t, 'Cum_V_res')
        sig_p.append(stars_p)

        p_cb_c = (cb_c / n_tot_b) if n_tot_b > 0 else 0.0
        p_cb_r = (cb_r / n_tot_b) if n_tot_b > 0 else 0.0
        p_ct_c = (ct_c / n_tot_t) if n_tot_t > 0 else 0.0
        p_ct_r = (ct_r / n_tot_t) if n_tot_t > 0 else 0.0

        crack_base_t.append(p_cb_c); rev_base_t.append(p_cb_r); total_base_t.append(p_cb_c + p_cb_r)
        cnt_crack_b.append(cb_c); cnt_rev_b.append(cb_r); cnt_tot_b.append(n_tot_b)
        
        crack_t4t_t.append(p_ct_c); rev_t4t_t.append(p_ct_r); total_t4t_t.append(p_ct_c + p_ct_r)
        cnt_crack_t.append(ct_c); cnt_rev_t.append(ct_r); cnt_tot_t.append(n_tot_t)
        
        _, stars_t = compute_paired_significance(sub_base, sub_t4t, 'Is_Terminated')
        sig_t.append(stars_t)

    err_base_p = [[min(m, s) for m, s in zip(means_base_p, stds_base_p)], stds_base_p]
    err_t4t_p = [[min(m, s) for m, s in zip(means_t4t_p, stds_t4t_p)], stds_t4t_p]

    rects1_p = ax1.bar(x_indices - bar_width/2, means_base_p, bar_width, yerr=err_base_p, capsize=4, color='#2b5c8f', edgecolor='black', linewidth=0.7, label="Base Simulation")
    rects2_p = ax1.bar(x_indices + bar_width/2, means_t4t_p, bar_width, yerr=err_t4t_p, capsize=4, color='#e05d44', edgecolor='black', linewidth=0.7, label="T4T Simulation")

    for rect, val in zip(rects1_p, means_base_p):
        if val > 0:
            ax1.text(rect.get_x() + rect.get_width()/2., val / 2., f"{val:.0f}",
                     ha='center', va='center', color='white', fontweight='bold', fontsize=8.5, rotation=90)
    for rect, val in zip(rects2_p, means_t4t_p):
        if val > 0:
            ax1.text(rect.get_x() + rect.get_width()/2., val / 2., f"{val:.0f}",
                     ha='center', va='center', color='white', fontweight='bold', fontsize=8.5, rotation=90)

    for x_i, mb, mt, sb, st, stars in zip(x_indices, means_base_p, means_t4t_p, stds_base_p, stds_t4t_p, sig_p):
        y_pos = max(mb + sb, mt + st) * 1.03 + 20
        font_color, font_weight = ('#d9534f', 'bold') if stars != 'ns' else ('#666666', 'normal')
        ax1.text(x_i, y_pos, stars, ha='center', va='bottom', fontsize=11, color=font_color, fontweight=font_weight)

    ax1.set_ylim(0, 8000)
    ax1.set_title('(A) Prosperity Scores (Mean Cum_V_res)', fontweight='bold', fontsize=11, pad=10)
    ax1.set_xticks(x_indices)
    ax1.set_xticklabels(pair_labels, fontweight='bold', fontsize=9.5)
    ax1.set_ylabel('Mean Cum_V_res', fontweight='bold', fontsize=10)
    ax1.grid(axis='y', linestyle='--', alpha=0.5)
    ax1.legend(loc='upper right', frameon=True, facecolor='white', framealpha=0.9, fontsize=8.0)

    x_base, x_t4t = x_indices - bar_width/2, x_indices + bar_width/2
    ax2.bar(x_base, crack_base_t, bar_width, label="Base Simulation (Crackdown)", color=color_base_crack, edgecolor='black', linewidth=0.7)
    ax2.bar(x_base, rev_base_t, bar_width, bottom=crack_base_t, label="Base Simulation (Revolution)", color=color_base_rev, edgecolor='black', linewidth=0.7)
    ax2.bar(x_t4t, crack_t4t_t, bar_width, label="T4T Simulation (Crackdown)", color=color_t4t_crack, edgecolor='black', linewidth=0.7)
    ax2.bar(x_t4t, rev_t4t_t, bar_width, bottom=crack_t4t_t, label="T4T Simulation (Revolution)", color=color_t4t_rev, edgecolor='black', linewidth=0.7)

    for i in range(n_pairs):
        # Base Bar Labels
        pc_b, nc_b = crack_base_t[i], cnt_crack_b[i]
        pr_b, nr_b = rev_base_t[i], cnt_rev_b[i]
        pt_b, nt_b = total_base_t[i], cnt_tot_b[i]

        if pc_b >= 0.04:
            ax2.text(x_base[i], pc_b / 2.0, f"{pc_b*100:.0f}%\n({nc_b})",
                     ha='center', va='center', color='white', fontweight='bold', fontsize=6.5)
        if pr_b >= 0.04:
            ax2.text(x_base[i], pc_b + pr_b / 2.0, f"{pr_b*100:.0f}%\n({nr_b})",
                     ha='center', va='center', color='#111111', fontweight='bold', fontsize=6.5)
        if pt_b > 0:
            ax2.text(x_base[i], pt_b + 0.01, f"{pt_b*100:.1f}%\n(n={nt_b})",
                     ha='center', va='bottom', color='black', fontweight='bold', fontsize=7.5)

        # T4T Bar Labels
        pc_t, nc_t = crack_t4t_t[i], cnt_crack_t[i]
        pr_t, nr_t = rev_t4t_t[i], cnt_rev_t[i]
        pt_t, nt_t = total_t4t_t[i], cnt_tot_t[i]

        if pc_t >= 0.04:
            ax2.text(x_t4t[i], pc_t / 2.0, f"{pc_t*100:.0f}%\n({nc_t})",
                     ha='center', va='center', color='white', fontweight='bold', fontsize=6.5)
        if pr_t >= 0.04:
            ax2.text(x_t4t[i], pc_t + pr_t / 2.0, f"{pr_t*100:.0f}%\n({nr_t})",
                     ha='center', va='center', color='#111111', fontweight='bold', fontsize=6.5)
        if pt_t > 0:
            ax2.text(x_t4t[i], pt_t + 0.01, f"{pt_t*100:.1f}%\n(n={nt_t})",
                     ha='center', va='bottom', color='black', fontweight='bold', fontsize=7.5)

        max_h = max(total_base_t[i], total_t4t_t[i])
        stars = sig_t[i]
        if max_h > 0:
            font_color, font_weight = ('#d9534f', 'bold') if stars != 'ns' else ('#666666', 'normal')
            ax2.text(x_indices[i], max_h + 0.07, stars, ha='center', va='bottom', fontsize=11, color=font_color, fontweight=font_weight)

    max_bar_val = max(max(total_base_t), max(total_t4t_t)) if total_base_t and total_t4t_t else 0.5
    ax2.set_ylim(0, max(max_bar_val * 1.35, 0.15))
    ax2.set_title('(B) Termination Proportions (Crackdown + Revolution)', fontweight='bold', fontsize=11, pad=10)
    ax2.set_xticks(x_indices)
    ax2.set_xticklabels(pair_labels, fontweight='bold', fontsize=9.5)
    ax2.set_ylabel('Termination Proportion (0.0 - 1.0)', fontweight='bold', fontsize=10)
    ax2.grid(axis='y', linestyle='--', alpha=0.5)
    ax2.legend(loc='upper left', frameon=True, facecolor='white', framealpha=0.9, fontsize=8.0)

    plt.suptitle(f'Archetype Paired Analysis: {archetype_name}', fontweight='bold', fontsize=13, y=1.05)
    plt.tight_layout()

    file_path = os.path.join(output_folder, f'dual_canvas_archetype_{archetype_name.lower()}.png')
    plt.savefig(file_path, dpi=300, bbox_inches='tight')
    plt.close()
    return file_path

def generate_graphs(df, target_dir):
    os.makedirs(target_dir, exist_ok=True)
    images = []
    sns.set_theme(style="whitegrid")
    
    print("Generating Dual-Metric Canvas Plots (Base vs TFT)...")
    for sim_base, sim_t4t, pair_label in SIM_PAIRS:
        p = plot_sim_pair_dual_canvas(df, sim_base, sim_t4t, pair_label, target_dir)
        images.append(p)
        
    print("Generating Vertically Stacked Cross-Archetype Overview...")
    p_cross = plot_cross_archetype_dual_canvas(df, target_dir)
    images.append(p_cross)
    
    print("Generating Individual Archetype Dual Canvases...")
    for arch in STANDARD_ARCHETYPE_ORDER:
        if arch in df['Archetype'].unique() and arch != 'Control':
            p_arch = plot_individual_archetype_dual_canvas(df, arch, target_dir)
            images.append(p_arch)

    return images

# ==========================================
# 5. REPORT & ARCHIVE GENERATION
# ==========================================
def generate_report(images, target_dir, df):
    doc = docx.Document()
    
    if 'Normal' in doc.styles:
        style_normal = doc.styles['Normal']
        style_normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        style_normal.font.name = 'Candara'
        
    for style_name in ['Title', 'Heading 1', 'Heading 2', 'Heading 3', 'Heading 4', 'Heading 5', 'Heading 6']:
        if style_name in doc.styles:
            h_style = doc.styles[style_name]
            h_style.font.color.rgb = RGBColor(0, 0, 0)
            h_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
            
    def add_centered_image(doc_obj, img_path, width):
        p = doc_obj.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run()
        r.add_picture(img_path, width=width)

    doc.add_heading('Ruler vs Ruled: Monte Carlo Simulation Report', 0)
    doc.add_paragraph("Ruler vs Ruled model created by the Pragmatic Realist (find me on Substack).")
    
    doc.add_heading('1. Introduction & Game Overview', level=1)
    doc.add_paragraph("This report presents the findings from an advanced Monte Carlo simulation of the 'Ruler vs. Ruled' dynamic. The model is structured as an escalating, multi-turn game of Chicken, where two actors (the Ruler and the Ruled) interact over a maximum of 1,000 steps per simulation.")
    doc.add_paragraph("In this environment, the 'Perceived Value of Society' (V_res) acts as a dynamic, shared resource representing general prosperity. During each step, both actors independently choose to act 'Hawkish' (aggressive/escalatory) or 'Dovish' (accommodating/de-escalatory). Hawkish actions extract a higher cost from the societal pool. Early termination of the simulation is triggered when systemic costs exhaust the total available value (PERCEIVED_VALUE = 10.0), indicating a structural societal failure (Revolution or Crackdown based on the distribution of pain).")
    
    doc.add_heading('2. Simulation Setup & Frameworks', level=1)
    doc.add_paragraph("The simulation sweeps across 8 distinct architectural frameworks to isolate the effects of differing strategic baselines and environmental sensitivities:")
    doc.add_paragraph("1. Simulation 1 (Stochastic Baseline): Completely random probability allocations at every step to map pure systemic noise.")
    doc.add_paragraph("2. Simulation 2 (Blind Control): Fixed ideological quadrants (DD, DH, HD, HH, RR) established at the start. These agents are perfectly oblivious to the decay of the societal resource pool.")
    doc.add_paragraph("3. Simulation 3 (Neutral Players): Random base hawkishness [0, 1] penalized by conscious awareness of structural decay (Aversion Penalties).")
    doc.add_paragraph("4. Simulation 4 (Neutral TFT): Players reactively copy the other's previous move. (If opponent played Dove, P_Hawk drops to [0, 0.5). If opponent played Hawk, P_Hawk jumps to (0.5, 1]).")
    doc.add_paragraph("5. Simulation 5 (Aggressive Ruler): Ruler has a fixed base hawkishness >0.5; Ruled remains neutral [0, 1].")
    doc.add_paragraph("6. Simulation 6 (Aggressive Ruler TFT): Ruler remains strictly Aggressive (>0.5); Ruled drops the neutral posture and adopts Tit-for-Tat copying.")
    doc.add_paragraph("7. Simulation 7 (Aggressive Ruled): Ruler is neutral; Ruled has a fixed base hawkishness >0.5.")
    doc.add_paragraph("8. Simulation 8 (Aggressive Ruled TFT): Ruler adopts TFT copying; Ruled remains strictly Aggressive (>0.5).")
    doc.add_paragraph("9. Simulation 9 (Blind TFT Control): A repeat of Simulation 2, but with both players using a Tit-for-Tat strategy, without responding to societal decay.")
    
    doc.add_heading('Strategic National Archetypes (Conscious Decay)', level=2)
    doc.add_paragraph("For simulations incorporating environmental decay awareness (Sims 3-8), we define five explicit aversion reactions:")
    doc.add_paragraph("1. Equals: Both Ruler and Ruled de-escalate collaboratively as resource pools drop (Linear 1:1 scaling).")
    doc.add_paragraph("2. Cowards: Symmetrically risk-averse, capitulating immediately into extreme dovish behavior under pressure (Quadratic scaling).")
    doc.add_paragraph("3. Fools: Symmetrically aggressive, resisting compromise and triggering stalemates (Sub-linear square-root scaling).")
    doc.add_paragraph("4. Tyrants: Hawkish Ruler (cushioned decay) and dovish Ruled (linear decay).")
    doc.add_paragraph("5. Brinksmen: Hawkish Ruled (cushioned decay) and dovish Ruler (linear decay).")

    
    doc.add_heading('3. State Failure Probability Analysis', level=1)
    doc.add_paragraph("This section details the exact probabilities of early state failure (and the corresponding triggers) across the strategic national groups. Each archetype is presented in its own table, with the blind Sim 2 Control and Sim 9 Control TFT included in every table as the comparative baselines.")

    archetypes = ['Equals', 'Cowards', 'Fools', 'Brinksmen', 'Tyrants']
    df_control_2 = df[df['Sim_Type'] == 2]
    df_control_9 = df[df['Sim_Type'] == 9]
    
    def get_failure_stats(sub_df):
        if not sub_df.empty:
            total_n = sub_df['Total_Count'].sum()
            survived_count = sub_df['Survived_Count'].sum()
            rev_count = sub_df['Revolution_Count'].sum()
            crack_count = sub_df['Crackdown_Count'].sum()
            if total_n > 0:
                return (total_n - survived_count) / total_n, rev_count / total_n, crack_count / total_n
        return 0.0, 0.0, 0.0

    ctrl_2_fail, ctrl_2_rev, ctrl_2_crack = get_failure_stats(df_control_2)
    ctrl_9_fail, ctrl_9_rev, ctrl_9_crack = get_failure_stats(df_control_9)
    
    for arch in archetypes:
        doc.add_heading(f"Archetype: {arch}", level=2)
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Light Shading Accent 1'
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text, hdr_cells[1].text = 'Configuration', 'Total Failure'
        hdr_cells[2].text, hdr_cells[3].text = 'Revolution (Ruler Lost)', 'Crackdown (Ruled Lost)'
        
        row_cells = table.add_row().cells
        row_cells[0].text = 'Sim 2 Control'
        row_cells[1].text = f"{ctrl_2_fail:.2%}"
        row_cells[2].text = f"{ctrl_2_rev:.2%}"
        row_cells[3].text = f"{ctrl_2_crack:.2%}"
        
        row_cells = table.add_row().cells
        row_cells[0].text = 'Sim 9 Control TFT'
        row_cells[1].text = f"{ctrl_9_fail:.2%}"
        row_cells[2].text = f"{ctrl_9_rev:.2%}"
        row_cells[3].text = f"{ctrl_9_crack:.2%}"
        
        for sim_type in range(3, 9):
            sname = {3:'Neutral', 4:'Neutral TFT', 5:'Agg Ruler', 6:'Agg Ruler TFT', 7:'Agg Ruled', 8:'Agg Ruled TFT'}[sim_type]
            sub_df = df[(df['Sim_Type'] == sim_type) & (df['Archetype'] == arch)]
            p_fail, p_rev, p_crack = get_failure_stats(sub_df)
            
            row_cells = table.add_row().cells
            row_cells[0].text = f"Sim {sim_type} {sname}"
            row_cells[1].text = f"{p_fail:.2%}"
            row_cells[2].text = f"{p_rev:.2%}"
            row_cells[3].text = f"{p_crack:.2%}"

    doc.add_heading('4. Analytical Visualizations', level=1)
    for img_path in images:
        add_centered_image(doc, img_path, Inches(6.5))
    
    doc.add_heading('5. Conclusion & Core Insights', level=1)
    doc.add_paragraph("Based on the comprehensive multi-framework analysis generated by this simulation run, several profound socio-structural dynamics emerge:")

    doc.add_paragraph("1. The Stabilizing Power of Tit-for-Tat (TFT): Across nearly all aggressive and unyielding archetypes (most notably the Fools and Tyrants), the introduction of a reactive, copy-based strategy dramatically curtails systemic collapse. By structurally guaranteeing retaliation against aggressive overreach, TFT functions as a powerful institutional governor, breaking runaway escalatory cycles before they drain the societal resource pool.")
    doc.add_paragraph("2. The Lethality of Strategic Blindness: The Sim 2 Control baseline operates with fixed ideological parameters and zero internal awareness of the wasting systemic value (V_res). This configuration consistently yields the highest rates of catastrophic collapse, which results in escalation spirals when Tit-for-Tat is introduced. This suggests that structural ignorance or political blindness to environmental/systemic decay is inherently more fatal to a society than explicit aggression. If neither side are willing to, or capable of reacting to the decline of society, because of foolishness or propaganda, the probability os systemic collapse is maximised.")
    doc.add_paragraph("3. The Preservation Efficiency of Symmetrical De-escalation: The 'Cowards' (highly risk-averse) and 'Equals' (cooperative) archetypes consistently demonstrate near-perfect survival rates and maximize the prosperity metric. In a multi-round environment characterized by mutual destruction (the Hawk-Dove dynamic), rapid, symmetrical capitulation in response to resource scarcity effectively insulates the collective wealth from structural friction.")

    doc.add_paragraph("")
    doc.add_paragraph("The Ruler vs Ruled Monte Carlo Simulation was created by the Pragmatic Realist: https://thepragmaticrealist.substack.com.")


    report_path = os.path.join(target_dir, FILE_REPORT)
    doc.save(report_path)
    return report_path

def create_archive(df_hist, report_path, plots_dir):
    counter = 1
    while os.path.exists(f"{OUTPUT_PREFIX}_{counter}.zip"):
        counter += 1
    zip_filename = f"{OUTPUT_PREFIX}_{counter}.zip"

    print(f"\rPreparing optimized export data sheets... ")

    summary_file = os.path.join(plots_dir, FILE_SUMMARY)
    history_file = os.path.join(plots_dir, FILE_HISTORY)
    feat_file = os.path.join(plots_dir, FILE_FEATURES)
    params_txt = os.path.join(plots_dir, FILE_PARAMS)
    
    if SORT_DATA:
        print("\nSorting huge output files... (Requires extensive RAM!)")
        for fpath in [summary_file, feat_file]:
            if os.path.exists(fpath):
                try:
                    df_sort = pd.read_csv(fpath)
                    df_sort.sort_values(by=['Sim_Type', 'Archetype', 'Case'], inplace=True)
                    df_sort.to_csv(fpath, index=False)
                    del df_sort
                except Exception as e:
                    print(f"Warning: Could not sort data in memory: {e}")
    
    if not df_hist.empty:
        print("Preparing History Spread Sheet...")
        df_hist_str = df_hist.applymap(lambda x: str(x) if pd.notna(x) else "")
        df_hist_str.to_csv(history_file)
    else:
        pd.DataFrame({"Status": ["History logging disabled."]}).to_csv(history_file, index=False)

    print("Preparing 'params.txt'...")
    with open(params_txt, "w") as f:
        f.write(f"PERCEIVED_VALUE = {PERCEIVED_VALUE}\n")
        f.write(f"MAX_COST_PER_STEP = {MAX_COST_PER_STEP}\n")
        f.write(f"MAX_STEPS = {MAX_STEPS}\n")
        f.write(f"N_REPS_PER_CONFIG = {N_REPS}\n")

    print("Preparing Graph Folder...")
    files_to_pack = [(report_path, FILE_REPORT), (summary_file, FILE_SUMMARY), (history_file, FILE_HISTORY), (feat_file, FILE_FEATURES), (params_txt, FILE_PARAMS)]
    for f in os.listdir(plots_dir):
        if f.endswith(".png"): files_to_pack.append((os.path.join(plots_dir, f), os.path.join("Graphs", f)))

    total_pack_files = len(files_to_pack)
    print(f"Packaging {total_pack_files} files directly into {zip_filename}...")
    start_pack_time = time.time()
    
    with zipfile.ZipFile(zip_filename, 'w') as zipf:
        for idx, (filepath, arcname) in enumerate(files_to_pack):
            zipf.write(filepath, arcname=arcname)
            completed = idx + 1
            if completed % max(1, total_pack_files // 10) == 0 or completed == total_pack_files:
                elapsed = time.time() - start_pack_time
                speed = completed / elapsed if elapsed > 0 else 0
                bar_length = 30
                filled_length = int(round(bar_length * completed / total_pack_files))
                bar_chars = '#' * filled_length + '-' * (bar_length - filled_length)
                percentage = (completed / total_pack_files) * 100
                print(f"\rPackaging ZIP Archive: [{bar_chars}] {completed}/{total_pack_files} ({percentage:.1f}%) | Speed: {speed:.1f} files/s", end="", flush=True)
            
    print("\nCleaning up intermediate workspace files...")
    shutil.rmtree(plots_dir)
    print(f"Archive successfully generated: {zip_filename}")

# ==========================================
# 6. MAIN EXECUTION ROUTINE
# ==========================================
if __name__ == "__main__":
    if os.path.exists(TEMP_DIR):
        shutil.rmtree(TEMP_DIR)
    os.makedirs(TEMP_DIR)
    
    init_persistent_workers()
    try:
        df_agg, df_hist = run_all_simulations()
        images = generate_graphs(df_agg, TEMP_DIR)
        report_path = generate_report(images, TEMP_DIR, df_agg)
        create_archive(df_hist, report_path, TEMP_DIR) # Can be commented out to save disc space
    finally:
        shutdown_persistent_workers()

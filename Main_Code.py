from __future__ import print_function
from cmath import inf
from docplex.mp.model import Model
from docplex.util.environment import get_environment
import win32com.client as win32
import pandas as pd
import os
import matplotlib.pyplot as plt
import numpy as np
from Result_Describe import writing_results

excel = win32.Dispatch("Excel.Application")
wb1 = excel.Workbooks.Open(os.getcwd()+"\\Data\\robust model_data.xlsx")
Price_DA = wb1.Sheets("Price_DA")              # Day-ahead prices
Price_DARS = wb1.Sheets("Price_DARS")          # Day-ahead reserve prices (Reserve prices)
Price_RT = wb1.Sheets("Price_RT")              # Real time prices (Up regulation prices)
Price_RTRS = wb1.Sheets("Price_RTRS")          # Real time reserve prices (Down regulation prices)
Expected_P_RT_WPR = wb1.Sheets("Expected_P_RT_WPR")  # Expected generation of renewable energy in the real-time operation

### Parameters
time_dim = 24     # Time
min_dim = 12      # ex) 5 minute x 12 = 1 hour (j)
del_S = 1/min_dim # Duration of intra-hourly interval ex) 5min = 1/12(h) 
BESS_dim = 2      # Num of BESS (s)
WPR_dim = 1       # Num of wind power (r)
Marginal_cost_CH = [1,0.8]    # Marginal cost of energy storage in charging modes
Marginal_cost_DCH = [1,0.8]   # Marginal cost of energy storage in discharging modes
Marginal_cost_WPR = [3]       # Marginal cost of renewable energy resources
Ramp_rate_WPR = 3             # Ramp-rate of renewable energy resources
E_min_BESS = [0,0]            # Minimum energy of energy storage
E_max_BESS = [30,18]          # Maximum energy of energy storage
P_max_BESS = [5,3]            # Maximum power of energy storage
P_min_BESS = [0,0]            # Minimum power of energy storage
Ramp_rate_BESS = [5,3]        # Ramp-rate of energy storage

Robust_percents = [0, 0.2, 0.4]   # Robust percent (0~1) / (Delta R, Variation interval in robust optimization)
contri_reg_percents = [0, 0.2, 0.4, 0.6, 0.8, 1] # Gamma, Maximum Contribution ratio of providing deployed up and down power in the real-time operation

### 최적화 파트
def build_optimization_model(Robust_percent, contri_reg_percent, name='Robust_Optimization_Model'):
    mdl = Model(name=name)   # Model
    mdl.parameters.mip.tolerances.mipgap = 0.0001;   # Tolerance

    time = [t for t in range(1,time_dim+1)]    # (t) one dimension
    time_min = [(t,j) for t in range(1,time_dim + 1) for j in range(1,min_dim+1)]   # (t,j) two dimension
    time_n_BESS = [(t,j,s) for t in range(1,time_dim + 1) for j in range(1,min_dim+1) for s in range(1,BESS_dim+1)]   # (t,j,s) three dimension
    time_n_WPR = [(t,j,r) for t in range(1,time_dim + 1) for j in range(1,min_dim+1) for r in range(1,WPR_dim+1)]     # (t,j,r) three dimension

    ### Continous Variable
    #Day-ahead
    P_DA_S = mdl.continuous_var_dict(time, lb=0, ub=inf, name="P-DA-S")   # Selling bids in the day-ahead market
    P_DA_B = mdl.continuous_var_dict(time, lb=0, ub=inf, name="P-DA-B")   # Buying bids in the day-ahead market
    P_RS = mdl.continuous_var_dict(time, lb=0, ub=inf, name="P-RS")       # Reserve bids in the day-ahead regulation market
    
    P_Up = mdl.continuous_var_dict(time_min, lb=0, ub=inf, name="P-Up")       # Deployed up power from the reserve power in the real-time operation (Deployed power in the up-regulation services)
    P_Down = mdl.continuous_var_dict(time_min, lb=0, ub=inf, name="P-Down")   # Deployed down power from the reserve power in the real-time operation (Deployed power in the down-regulation services)

    P_DA_CH = mdl.continuous_var_dict(time_n_BESS, lb=0, ub=inf, name="P-DA-CH")      # Day-ahead scheduling of energy storage in charging modes in the day-ahead scheduling
    P_DA_DCH = mdl.continuous_var_dict(time_n_BESS, lb=0, ub=inf, name="P-DA-DCH")    # Day-ahead scheduling of energy storage in discharging modes in the day-ahead scheduling
    P_DA_Energy = mdl.continuous_var_dict(time_n_WPR, lb=0, ub=inf, name="P-DA-Energy")  # Day-ahead scheduling of renewable energy resources in the day-ahead scheduling
    
    P_Up_CH = mdl.continuous_var_dict(time_n_BESS, lb=0, ub=inf, name="P-Up-CH")     # Deployed up power from the reserve power of energy storage in charging mode in the real-time operation
    P_Up_DCH = mdl.continuous_var_dict(time_n_BESS, lb=0, ub=inf, name="P-Up-DCH")   # Deployed up power from the reserve power of energy storage in discharging mode in the real-time operation
    P_Up_RES = mdl.continuous_var_dict(time_n_WPR, lb=0, ub=inf, name="P-Up-RES")    # Deployed up power from the reserve power of renewable energy resources in the real-time operation

    P_Down_CH = mdl.continuous_var_dict(time_n_BESS, lb=0, ub=inf, name="P-Down-CH")    # Deployed down power from the reserve power of energy storage in charging mode in the real-time operation
    P_Down_DCH = mdl.continuous_var_dict(time_n_BESS, lb=0, ub=inf, name="P-Down-DCH")  # Deployed down power from the reserve power of energy storage in discharging mode in the real-time operation
    P_Down_RES = mdl.continuous_var_dict(time_n_WPR, lb=0, ub=inf, name="P-Down-RES")   # Deployed down power from the reserve power of renewable energy resources in the real-time operation

    P_RS_CH = mdl.continuous_var_dict(time_n_BESS, lb=0, ub=inf, name="P-RS-CH")      # Reserve scheduling of energy storage in charging modes
    P_RS_DCH = mdl.continuous_var_dict(time_n_BESS, lb=0, ub=inf, name="P-RS-DCH")    # Reserve scheduling of energy storage in discharging modes
    P_RS_RES = mdl.continuous_var_dict(time_n_WPR, lb=0, ub=inf, name="P-RS-RES")     # Reserve scheduling of renewable energy resources

    P_RES_imb = mdl.continuous_var_dict(time_n_WPR, lb=0, ub=inf, name="P-RES-imb")    # Imbalance power between the day-ahead scheduling and the real-time operation of renewable energy resources
    E_BESS = mdl.continuous_var_dict(time_n_BESS, lb=0, ub=inf, name="E-BESS")        # Energy level of energy storage
    P_RT_WPR = mdl.continuous_var_dict(time_n_WPR, lb=0, ub=inf, name="P-RT-WPR")         # Real-time power of renewable energy resources
    
    ### Functions

    AV_RO = mdl.continuous_var(lb=0, ub=inf, name="AV-RO")          # Auxiliary variable of robust optimization
    I_t = mdl.continuous_var_dict(time, lb=0, ub=inf, name="B-t")   # Income function of owner
    C_t = mdl.continuous_var_dict(time, lb=0, ub=inf, name="C-t")   # Cost function of owner

    ### Binary Variable
    D_Char = mdl.binary_var_dict(time_n_BESS, name="D-Char-DA")      # Charging binary variable of energy storage (알파)
    D_Dchar = mdl.binary_var_dict(time_n_BESS, name="D-DChar-DA")    # Discharging binary variable of energy storage (베타)
    D_WPR = mdl.binary_var_dict(time_n_WPR, name="D-WPR")            # Commitment status binary variable of renewable energy resources
    
    ### Robust Optimization - Equation(57)
    mdl.maximize(mdl.sum(Price_DA.Cells(t+1,2).Value * mdl.sum(del_S * (mdl.sum(P_DA_DCH[(t,j,s)] - P_DA_CH[(t,j,s)] for s in range(1,BESS_dim+1)) + mdl.sum(P_DA_Energy [(t,j,r)] for r in range(1,WPR_dim+1))) for j in range(1,min_dim+1))
                         + Price_DARS.Cells(t+1,2).Value * mdl.sum(del_S * ((mdl.sum(P_RS_CH[(t,j,s)] + P_RS_DCH[(t,j,s)] for s in range(1,BESS_dim+1)) + mdl.sum(P_RS_RES[(t,j,r)] for r in range(1,WPR_dim+1)))) for j in range(1,min_dim+1))
                         - mdl.sum((mdl.sum(Marginal_cost_DCH[s-1] * del_S * P_DA_DCH[(t,j,s)] + Marginal_cost_CH[s-1] * del_S * P_DA_CH[(t,j,s)] for s in range(1,BESS_dim+1)) + mdl.sum(Marginal_cost_WPR[r-1] * del_S * P_DA_Energy [(t,j,r)] for r in range(1,WPR_dim+1))) for j in range(1,min_dim+1))
                         for t in range(1,time_dim+1))
                 + AV_RO)

    # Robust Optizimation Equation(58)
    mdl.add_constraint(AV_RO <= mdl.sum(Price_RT.Cells(t+1,j+1).Value * del_S * (mdl.sum(P_Up_DCH[(t,j,s)] - P_Up_CH[(t,j,s)] - P_Down_CH[(t,j,s)] + P_Down_DCH[(t,j,s)] for s in range(1,BESS_dim+1)) + mdl.sum(P_Up_RES[(t,j,r)] - P_Down_RES[(t,j,r)] - P_RES_imb[(t,j,r)] for r in range(1,WPR_dim+1)))
                                        + Price_RTRS.Cells(t+1,j+1).Value * del_S * (mdl.sum(P_Down_CH[(t,j,s)] - P_Down_DCH[(t,j,s)] for s in range(1,BESS_dim+1)) + mdl.sum(P_Down_RES[(t,j,r)] for r in range(1,WPR_dim+1))) 
                                            - mdl.sum(Marginal_cost_DCH[s-1] * del_S * (P_Up_DCH[(t,j,s)] + P_Down_DCH[(t,j,s)]) + Marginal_cost_CH[s-1] * del_S * (P_Up_CH[(t,j,s)] + P_Down_CH[(t,j,s)]) for s in range(1,BESS_dim+1)) - mdl.sum(Marginal_cost_WPR[r-1] * del_S * P_Up_RES[(t,j,r)] for r in range(1,WPR_dim+1)) 
                                            for j in range(1,min_dim+1) for t in range(1,time_dim+1)))
    
    # Day-ahead bid Equation(4) ~ Equation(8)
    mdl.add_constraints(P_DA_S[t] == mdl.sum(mdl.sum(P_DA_DCH[(t,j,s)] for s in range(1,BESS_dim+1)) + mdl.sum(P_DA_Energy[(t,j,r)] for r in range(1,WPR_dim+1))) for j in range(1,2) for t in range(1,time_dim+1))  # Equation(4)

    mdl.add_constraints(P_DA_B[t] == mdl.sum(P_DA_CH[(t,j,s)] for s in range(1,BESS_dim+1)) for j in range(1,2) for t in range(1,time_dim+1))  # Equation(5)    
    
    mdl.add_constraints(P_DA_DCH[(t,j,s)] == P_DA_DCH[(t,J,s)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for J in range(1,min_dim+1) for s in range(1,BESS_dim+1))  # Equation(6)

    mdl.add_constraints(P_DA_CH[(t,j,s)] == P_DA_CH[(t,J,s)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for J in range(1,min_dim+1) for s in range(1,BESS_dim+1))    # Equation(7)

    mdl.add_constraints(P_DA_Energy[(t,j,r)] == P_DA_Energy [(t,J,r)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for J in range(1,min_dim+1) for r in range(1,WPR_dim+1))   # Equation(8)

    # Reserve bid Equation(9) ~ Equation(13)
    mdl.add_constraints(P_RS[t] == mdl.sum(mdl.sum(P_RS_CH[(t,j,s)] + P_RS_DCH[(t,j,s)] for s in range(1,BESS_dim+1)) + mdl.sum(P_RS_RES[(t,j,r)] for r in range(1,WPR_dim+1))) for j in range(1,2) for t in range(1,time_dim+1))  # Equation(9)
        
    mdl.add_constraints(P_RS_CH[(t,j,s)] == P_RS_CH[(t,J,s)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for J in range(1,min_dim+1) for s in range(1,BESS_dim+1))    # Equation(10)  

    mdl.add_constraints(P_RS_DCH[(t,j,s)] == P_RS_DCH[(t,J,s)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for J in range(1,min_dim+1) for s in range(1,BESS_dim+1))  # Equation(11)                                          

    mdl.add_constraints(P_RS_RES[(t,j,r)] == P_RS_RES[(t,J,r)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for J in range(1,min_dim+1) for r in range(1,WPR_dim+1))   # Equation(12)
       
    mdl.add_constraints(P_RS[t] <= mdl.sum(contri_reg_percent * P_max_BESS[s-1] for s in range(1,BESS_dim+1)) + mdl.sum(contri_reg_percent * P_DA_Energy[(t,min_dim,r)] for r in range(1,WPR_dim+1)) for t in range(1,time_dim+1)) # Equation(13), Equation(14)
    
    # Deployed up/down power from the reserve power in the real-time operation Equation(15) ~ (18)
    mdl.add_constraints(P_Up[(t,j)] <= P_RS[t] for t in range(1,time_dim+1) for j in range(1,min_dim+1))    # Equation(15)
    
    mdl.add_constraints(P_Down[(t,j)] <= P_RS[t] for t in range(1,time_dim+1) for j in range(1,min_dim+1))  # Equation(16)
    
    mdl.add_constraints(P_Up[(t,j)] == mdl.sum(mdl.sum(P_Up_DCH[(t,j,s)] - P_Up_CH[(t,j,s)] for s in range(1,BESS_dim+1)) + mdl.sum(P_Up_RES[(t,j,r)] for r in range(1,WPR_dim+1))) for j in range(1,min_dim+1) for t in range(1,time_dim+1))  # Equation(17)
    
    mdl.add_constraints(P_Down[(t,j)] == mdl.sum(mdl.sum(P_Down_CH[(t,j,s)] - P_Down_DCH[(t,j,s)] for s in range(1,BESS_dim+1)) + mdl.sum(P_Down_RES[(t,j,r)] for r in range(1,WPR_dim+1))) for j in range(1,min_dim+1) for t in range(1,time_dim+1))  # Equation(18)
                          
    ### Constarints of capacity - Equation(28) ~ Equation(31) / Equation(35) ~ Equation(40) / Equation(44) ~ Equation(45) / Equation(47)
    for t in range(1,time_dim+1):
        for j in range(1,min_dim+1):
            for s in range(1,BESS_dim+1):
                mdl.add_constraint(P_DA_CH[(t,j,s)] <= P_max_BESS[s-1] * D_Char[(t,j,s)])                         # Equation(28)
                mdl.add_constraint(P_min_BESS[s-1] * D_Char[(t,j,s)] <= P_DA_CH[(t,j,s)])                         # Equation(28)
                
                mdl.add_constraint(P_RS_CH[(t,j,s)] <= P_max_BESS[s-1] * D_Char[(t,j,s)] - P_DA_CH[(t,j,s)])     # Equation(29)
                mdl.add_constraint(P_min_BESS[s-1] <= P_RS_CH[(t,j,s)])                                          # Equation(29)
                
                mdl.add_constraint(P_DA_CH[(t,j,s)] + P_RS_CH[(t,j,s)] <= P_max_BESS[s-1] * D_Char[(t,j,s)])     # Equation(30)
                
                mdl.add_constraint(P_min_BESS[s-1] * D_Char[(t,j,s)] <= P_DA_CH[(t,j,s)] - P_RS_CH[(t,j,s)])     # Equation(31)
                
                mdl.add_constraint(P_Up_CH[(t,j,s)] <= P_RS_CH[(t,j,s)])                                         # Equation(35)
                
                mdl.add_constraint(P_Down_CH[(t,j,s)] <= P_RS_CH[(t,j,s)])                                       # Equation(36)   
                             
                mdl.add_constraint(P_DA_DCH[(t,j,s)] <= P_max_BESS[s-1] * D_Dchar[(t,j,s)])                      # Equation(37)
                mdl.add_constraint(P_min_BESS[s-1] * D_Dchar[(t,j,s)] <= P_DA_DCH[(t,j,s)])                      # Equation(37)
                
                mdl.add_constraint(P_RS_DCH[(t,j,s)] <= P_max_BESS[s-1] * D_Dchar[(t,j,s)] - P_DA_DCH[(t,j,s)])  # Equation(38)
                mdl.add_constraint(P_min_BESS[s-1] <= P_RS_DCH[(t,j,s)])                                         # Equation(38)

                mdl.add_constraint(P_DA_DCH[(t,j,s)] + P_RS_DCH[(t,j,s)] <= P_max_BESS[s-1] * D_Dchar[(t,j,s)])  # Equation(39)
                   
                mdl.add_constraint(P_min_BESS[s-1] * D_Dchar[(t,j,s)] <= P_DA_DCH[(t,j,s)] - P_RS_DCH[(t,j,s)])  # Equation(40)
                            
                mdl.add_constraint(P_Up_DCH[(t,j,s)] <= P_RS_DCH[(t,j,s)])                                       # Equation(44)
                
                mdl.add_constraint(P_Down_DCH[(t,j,s)] <= P_RS_DCH[(t,j,s)])                                     # Equation(45)
                        
                mdl.add_constraint(E_min_BESS[s-1] <= E_BESS[(t,j,s)])  # Equation(47)
                mdl.add_constraint(E_BESS[(t,j,s)] <= E_max_BESS[s-1])  # Equation(47)
                
    # Operation Constraints - Renewable Energy - Equation(19) ~ Equation(27)
    mdl.add_constraints(P_DA_Energy[(t,j,r)] <= P_RT_WPR[(t,j,r)]  for t in range(1,time_dim+1) for j in range(1,min_dim+1) for r in range(1,WPR_dim+1))  # Equation(19) 
    
    mdl.add_constraints(P_RS_RES[(t,j,r)] <= P_RT_WPR[(t,j,r)] - P_DA_Energy[(t,j,r)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for r in range(1,WPR_dim+1))   # Equation(20)

    mdl.add_constraints(P_DA_Energy[(t,j,r)] + P_RS_RES[(t,j,r)] <= P_RT_WPR[(t,j,r)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for r in range(1,WPR_dim+1))   # Equation(21)
    
    mdl.add_constraints(0 <= P_DA_Energy[(t,j,r)] - P_RS_RES[(t,j,r)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for r in range(1,WPR_dim+1))   # Equation(22)

    ##Equation(23)
    #t>=1, j>=2
    mdl.add_constraints(-1 * Ramp_rate_WPR <= P_DA_Energy[(t,j,r)] - P_DA_Energy[(t,j-1,r)] for t in range(1,time_dim+1) for j in range(2,min_dim+1) for r in range(1,WPR_dim+1))
    mdl.add_constraints(Ramp_rate_WPR >= P_DA_Energy[(t,j,r)] - P_DA_Energy[(t,j-1,r)] for t in range(1,time_dim+1) for j in range(2,min_dim+1) for r in range(1,WPR_dim+1))
    
    #t>=2 and j=1 
    mdl.add_constraints(-1 * Ramp_rate_WPR <= P_DA_Energy[(t,j,r)] - P_DA_Energy[(t-1,min_dim,r)] for t in range(2,time_dim+1) for j in range(1,2) for r in range(1,WPR_dim+1))
    mdl.add_constraints(Ramp_rate_WPR >= P_DA_Energy[(t,j,r)] - P_DA_Energy[(t-1,min_dim,r)] for t in range(2,time_dim+1) for j in range(1,2) for r in range(1,WPR_dim+1))
    
    #t=1 and j=1 
    mdl.add_constraints(-1 * Ramp_rate_WPR <= P_DA_Energy[(t,j,r)] for t in range(1,2) for j in range(1,2) for r in range(1,WPR_dim+1))
    mdl.add_constraints(Ramp_rate_WPR >= P_DA_Energy[(t,j,r)] for t in range(1,2) for j in range(1,2) for r in range(1,WPR_dim+1))
    
    ##Equation(24)
    #t>=1, j>=2
    mdl.add_constraints(Ramp_rate_WPR >= P_RS_RES[(t,j,r)] + P_RS_RES[(t,j-1,r)] for t in range(1,time_dim+1) for j in range(2,min_dim+1) for r in range(1,WPR_dim+1))
    
    #t>=2 and j=1 
    mdl.add_constraints(Ramp_rate_WPR >= P_RS_RES[(t,j,r)] + P_RS_RES[(t-1,min_dim,r)] for t in range(2,time_dim+1) for j in range(1,2) for r in range(1,WPR_dim+1))
    
    #t=1 and j=1 
    mdl.add_constraints(Ramp_rate_WPR >= P_RS_RES[(t,j,r)] for t in range(1,2) for j in range(1,2) for r in range(1,WPR_dim+1))
    
    ##Equation(25)
    #t>=1, j>=2
    mdl.add_constraints(-1 * Ramp_rate_WPR<= (P_DA_Energy[(t,j,r)] - P_DA_Energy[(t,j-1,r)]) + (P_RS_RES[(t,j,r)] + P_RS_RES[(t,j-1,r)]) for t in range(1,time_dim+1) for j in range(2,min_dim+1) for r in range(1,WPR_dim+1))
    mdl.add_constraints(Ramp_rate_WPR >= (P_DA_Energy[(t,j,r)] - P_DA_Energy[(t,j-1,r)]) + (P_RS_RES[(t,j,r)] + P_RS_RES[(t,j-1,r)]) for t in range(1,time_dim+1) for j in range(2,min_dim+1) for r in range(1,WPR_dim+1))
    
    #t>=2 and j=1 
    mdl.add_constraints(-1 * Ramp_rate_WPR <= (P_DA_Energy[(t,j,r)] - P_DA_Energy[(t-1,min_dim,r)] ) + (P_RS_RES[(t,j,r)] + P_RS_RES[(t-1,min_dim,r)]) for t in range(2,time_dim+1) for j in range(1,2) for r in range(1,WPR_dim+1))
    mdl.add_constraints(Ramp_rate_WPR >= (P_DA_Energy[(t,j,r)] - P_DA_Energy[(t-1,min_dim,r)]) + (P_RS_RES[(t,j,r)] + P_RS_RES[(t-1,min_dim,r)]) for t in range(2,time_dim+1) for j in range(1,2) for r in range(1,WPR_dim+1))
    
    #t=1 and j=1
    mdl.add_constraints(-1 * Ramp_rate_WPR <= (P_DA_Energy[(t,j,r)] + P_RS_RES[(t,j,r)]) for t in range(1,2) for j in range(1,2) for r in range(1,WPR_dim+1))
    mdl.add_constraints(Ramp_rate_WPR >= (P_DA_Energy[(t,j,r)] + P_RS_RES[(t,j,r)]) for t in range(1,2) for j in range(1,2) for r in range(1,WPR_dim+1))

    # Equation(26)
    mdl.add_constraints(0 <= P_Up_RES[(t,j,r)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for r in range(1,WPR_dim+1))
    
    mdl.add_constraints(P_Up_RES[(t,j,r)] <= P_RS_RES[(t,j,r)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for r in range(1,WPR_dim+1))

    # Equation(27)
    mdl.add_constraints(0 <= P_Down_RES[(t,j,r)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for r in range(1,WPR_dim+1))

    mdl.add_constraints(P_Down_RES[(t,j,r)] <= P_RS_RES[(t,j,r)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for r in range(1,WPR_dim+1))
       
    ### Constarints of ramp-rate (BESS)
    ##Equation(32)
    #t>=1, j>=2
    mdl.add_constraints(-1 * Ramp_rate_BESS[s-1] <= P_DA_CH[(t,j,s)] - P_DA_CH[(t,j-1,s)] for t in range(1,time_dim+1) for j in range(2,min_dim+1) for s in range(1,BESS_dim+1))
    mdl.add_constraints(Ramp_rate_BESS[s-1] >= P_DA_CH[(t,j,s)] - P_DA_CH[(t,j-1,s)] for t in range(1,time_dim+1) for j in range(2,min_dim+1) for s in range(1,BESS_dim+1))
    
    #t>=2 and j=1 
    mdl.add_constraints(-1 * Ramp_rate_BESS[s-1] <= P_DA_CH[(t,j,s)] - P_DA_CH[(t-1,min_dim,s)] for t in range(2,time_dim+1) for j in range(1,2) for s in range(1,BESS_dim+1))
    mdl.add_constraints(Ramp_rate_BESS[s-1] >= P_DA_CH[(t,j,s)] - P_DA_CH[(t-1,min_dim,s)] for t in range(2,time_dim+1) for j in range(1,2) for s in range(1,BESS_dim+1))
    
    #t=1 and j=1
    mdl.add_constraints(-1 * Ramp_rate_BESS[s-1] <= P_DA_CH[(t,j,s)] for t in range(1,2) for j in range(1,2) for s in range(1,BESS_dim+1))
    mdl.add_constraints(Ramp_rate_BESS[s-1] >= P_DA_CH[(t,j,s)] for t in range(1,2) for j in range(1,2) for s in range(1,BESS_dim+1))
    
    ##Equation(33)
    #t>=1 and j>=2
    mdl.add_constraints(Ramp_rate_BESS[s-1] >= P_RS_CH[(t,j,s)] + P_RS_CH[(t,j-1,s)] for t in range(1,time_dim+1) for j in range(2,min_dim+1) for s in range(1,BESS_dim+1))
    
    #t>=2 and j=1 
    mdl.add_constraints(Ramp_rate_BESS[s-1] >= P_RS_CH[(t,j,s)] + P_RS_CH[(t-1,min_dim,s)] for t in range(2,time_dim+1) for j in range(1,2) for s in range(1,BESS_dim+1))
    
    #t=1 and j=1
    mdl.add_constraints(Ramp_rate_BESS[s-1] >= P_RS_CH[(t,j,s)] for t in range(1,2) for j in range(1,2) for s in range(1,BESS_dim+1)) 

    ##Equation(34)
    #t>=1, j>=2
    mdl.add_constraints(-1 * Ramp_rate_BESS[s-1] <= (P_DA_CH[(t,j,s)] - P_DA_CH[(t,j-1,s)]) + (P_RS_CH[(t,j,s)] + P_RS_CH[(t,j-1,s)]) for t in range(1,time_dim+1) for j in range(2,min_dim+1) for s in range(1,BESS_dim+1))
    mdl.add_constraints(Ramp_rate_BESS[s-1] >= (P_DA_CH[(t,j,s)] - P_DA_CH[(t,j-1,s)]) + (P_RS_CH[(t,j,s)] + P_RS_CH[(t,j-1,s)]) for t in range(1,time_dim+1) for j in range(2,min_dim+1) for s in range(1,BESS_dim+1))
    
    #t>=2 and j=1 
    mdl.add_constraints(-1 * Ramp_rate_BESS[s-1] <= (P_DA_CH[(t,j,s)] - P_DA_CH[(t-1,min_dim,s)]) + (P_RS_CH[(t,j,s)] + P_RS_CH[(t-1,min_dim,s)]) for t in range(2,time_dim+1) for j in range(1,2) for s in range(1,BESS_dim+1))
    mdl.add_constraints(Ramp_rate_BESS[s-1] >= (P_DA_CH[(t,j,s)] - P_DA_CH[(t-1,min_dim,s)]) + (P_RS_CH[(t,j,s)]+ P_RS_CH[(t-1,min_dim,s)]) for t in range(2,time_dim+1) for j in range(1,2) for s in range(1,BESS_dim+1))
    
    #t=1 and j=1 
    mdl.add_constraints(-1 * Ramp_rate_BESS[s-1] <= (P_DA_CH[(t,j,s)] + P_RS_CH[(t,j,s)]) for t in range(1,2) for j in range(1,2) for s in range(1,BESS_dim+1))
    mdl.add_constraints(Ramp_rate_BESS[s-1] >= (P_DA_CH[(t,j,s)] + P_RS_CH[(t,j,s)]) for t in range(1,2) for j in range(1,2) for s in range(1,BESS_dim+1))
    
    ##Equation(41)
    #t>=1 and j>=2
    mdl.add_constraints(-1 * Ramp_rate_BESS[s-1] <= P_DA_DCH[(t,j,s)] - P_DA_DCH[(t,j-1,s)] for t in range(1,time_dim+1) for j in range(2,min_dim+1) for s in range(1,BESS_dim+1))
    mdl.add_constraints(Ramp_rate_BESS[s-1] >= P_DA_DCH[(t,j,s)] - P_DA_DCH[(t,j-1,s)] for t in range(1,time_dim+1) for j in range(2,min_dim+1) for s in range(1,BESS_dim+1))
    
    #t>=2 and j=1 
    mdl.add_constraints(-1 * Ramp_rate_BESS[s-1] <= P_DA_DCH[(t,j,s)] - P_DA_DCH[(t-1,min_dim,s)] for t in range(2,time_dim+1) for j in range(1,2) for s in range(1,BESS_dim+1))
    mdl.add_constraints(Ramp_rate_BESS[s-1] >= P_DA_DCH[(t,j,s)] - P_DA_DCH[(t-1,min_dim,s)] for t in range(2,time_dim+1) for j in range(1,2) for s in range(1,BESS_dim+1))
    
    #t=1 and j=1
    mdl.add_constraints(-1 * Ramp_rate_BESS[s-1] <= P_DA_DCH[(t,j,s)] for t in range(1,2) for j in range(1,2) for s in range(1,BESS_dim+1))
    mdl.add_constraints(Ramp_rate_BESS[s-1] >= P_DA_DCH[(t,j,s)] for t in range(1,2) for j in range(1,2) for s in range(1,BESS_dim+1))
       
    ##Equation(42)
    #t>=1 and j>=2
    mdl.add_constraints(Ramp_rate_BESS[s-1] >= P_RS_DCH[(t,j,s)] + P_RS_DCH[(t,j-1,s)] for t in range(1,time_dim+1) for j in range(2,min_dim+1) for s in range(1,BESS_dim+1))
    
    #t>=2 and j=1 
    mdl.add_constraints(Ramp_rate_BESS[s-1] >= P_RS_DCH[(t,j,s)] + P_RS_DCH[(t-1,min_dim,s)] for t in range(2,time_dim+1) for j in range(1,2) for s in range(1,BESS_dim+1))
    
    #t=1 and j=1 
    mdl.add_constraints(Ramp_rate_BESS[s-1] >= P_RS_DCH[(t,j,s)] for t in range(1,2) for j in range(1,2) for s in range(1,BESS_dim+1))
    
    ##Equation(43)
    #t>=1, j>=2 
    mdl.add_constraints(-1 * Ramp_rate_BESS[s-1] <= (P_DA_DCH[(t,j,s)] - P_DA_DCH[(t,j-1,s)] ) + (P_RS_DCH[(t,j,s)] + P_RS_DCH[(t,j-1,s)]) for t in range(1,time_dim+1) for j in range(2,min_dim+1) for s in range(1,BESS_dim+1))
    mdl.add_constraints(Ramp_rate_BESS[s-1] >= (P_DA_DCH[(t,j,s)] - P_DA_DCH[(t,j-1,s)] ) + (P_RS_DCH[(t,j,s)] + P_RS_DCH[(t,j-1,s)]) for t in range(1,time_dim+1) for j in range(2,min_dim+1) for s in range(1,BESS_dim+1))
    
    #t>=2 and j=1
    mdl.add_constraints(-1 * Ramp_rate_BESS[s-1] <= (P_DA_DCH[(t,j,s)] - P_DA_DCH[(t-1,min_dim,s)] ) + (P_RS_DCH[(t,j,s)] + P_RS_DCH[(t-1,min_dim,s)]) for t in range(2,time_dim+1) for j in range(1,2) for s in range(1,BESS_dim+1))
    mdl.add_constraints(Ramp_rate_BESS[s-1] >= (P_DA_DCH[(t,j,s)] - P_DA_DCH[(t-1,min_dim,s)] ) + (P_RS_DCH[(t,j,s)] + P_RS_DCH[(t-1,min_dim,s)]) for t in range(2,time_dim+1) for j in range(1,2) for s in range(1,BESS_dim+1))
    
    #t=1 and j=1
    mdl.add_constraints(-1 * Ramp_rate_BESS[s-1] <= (P_DA_DCH[(t,j,s)] + P_RS_DCH[(t,j,s)]) for t in range(1,2) for j in range(1,2) for s in range(1,BESS_dim+1))
    mdl.add_constraints(Ramp_rate_BESS[s-1] >= (P_DA_DCH[(t,j,s)] + P_RS_DCH[(t,j,s)]) for t in range(1,2) for j in range(1,2) for s in range(1,BESS_dim+1))

    ### Constarints of stored energy of BESS - Equation(46)
    ## t>=1, j>=2
    mdl.add_constraints(E_BESS[(t,j,s)] == E_BESS[(t,j-1,s)] + del_S * (P_DA_CH[(t,j,s)] - P_DA_DCH[(t,j,s)] + P_Down_CH[(t,j,s)] + P_Up_CH[(t,j,s)] - P_Up_DCH[(t,j,s)] - P_Down_DCH[(t,j,s)]) 
                       for t in range(1, time_dim+1) for j in range(2, min_dim+1) for s in range(1,BESS_dim+1))
    
    ## t>=2, j=1
    mdl.add_constraints(E_BESS[(t,j,s)] == E_BESS[(t-1,min_dim,s)] + del_S * (P_DA_CH[(t,j,s)] - P_DA_DCH[(t,j,s)] + P_Down_CH[(t,j,s)] + P_Up_CH[(t,j,s)] - P_Up_DCH[(t,j,s)] - P_Down_DCH[(t,j,s)]) 
                       for t in range(2, time_dim+1) for j in range(1, 2) for s in range(1,BESS_dim+1))
    
    ## t=1, j=1
    mdl.add_constraints(E_BESS[(t,j,s)] == E_max_BESS[s-1]/2 
                       for t in range(1, 2) for j in range(1, 2) for s in range(1,BESS_dim+1))
    
    ## t=T, j=Nj
    mdl.add_constraints(E_BESS[(t,j,s)] == E_max_BESS[s-1]/2 
                       for t in range(time_dim, time_dim+1) for j in range(min_dim, min_dim+1) for s in range(1,BESS_dim+1))
 
    ### Constarints of binary decision Variables - Equation(48) ~ Equation(50)
    mdl.add_constraints(D_Char[(t,j,s)] == D_Char[(t,J,s)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for J in range(1,min_dim+1) for s in range(1,BESS_dim+1))    # Equation(48)

    mdl.add_constraints(D_Dchar[(t,j,s)] == D_Dchar[(t,J,s)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for J in range(1,min_dim+1) for s in range(1,BESS_dim+1))  # Equation(49)

    mdl.add_constraints(0 <= D_Char[(t,j,s)] + D_Dchar[(t,j,s)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for s in range(1,BESS_dim+1))  # Equation(50)

    mdl.add_constraints(D_Char[(t,j,s)] + D_Dchar[(t,j,s)] <= 1 for t in range(1,time_dim+1) for j in range(1,min_dim+1) for s in range(1,BESS_dim+1))  # Equation(50)
    
    mdl.add_constraints(D_WPR[(t,j,r)] == D_WPR[(t,J,r)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for J in range(1,min_dim+1) for r in range(1,WPR_dim+1))       # Equation(56) 
          
    ### Constarints of  Imbalance power from renewable energy - Equation(51) ~ Equation(52)
    ##Equation(51)
    mdl.add_constraints(P_RES_imb[(t,j,r)] == P_RT_WPR[(t,j,r)] - (P_DA_Energy[(t,j,r)] + P_Up_RES[(t,j,r)] - P_Down_RES[(t,j,r)]) 
                        for t in range(1,time_dim+1) for j in range(1,min_dim+1) for r in range(1,WPR_dim+1))
    
    ##Equation(52)
    mdl.add_constraints(P_RES_imb[(t,j,r)] <= P_RT_WPR[(t,j,r)] 
                        for t in range(1,time_dim+1) for j in range(1,min_dim+1) for r in range(1,WPR_dim+1))

                        
    ### Constraints of uncertain parameters-  Equation(53) ~ Equation(55)
    if Robust_percent != 0:
        mdl.add_constraints((-Robust_percent) * (mdl.sum(del_S * contri_reg_percent * P_max_BESS[s-1] for s in range(1,BESS_dim+1)) + mdl.sum(del_S * contri_reg_percent * P_DA_Energy[(t,j,r)] for r in range(1,WPR_dim+1))) <= P_Up[(t,j)] - mdl.sum(del_S * contri_reg_percent * P_max_BESS[s-1] for s in range(1,BESS_dim+1)) - mdl.sum(del_S * contri_reg_percent * P_DA_Energy[(t,j,r)] for r in range(1,WPR_dim+1)) for t in range(1,time_dim+1) for j in range(1,min_dim+1)) # Equation(53)

        mdl.add_constraints((Robust_percent) * (mdl.sum(del_S * contri_reg_percent * P_max_BESS[s-1] for s in range(1,BESS_dim+1)) + mdl.sum(del_S * contri_reg_percent * P_DA_Energy[(t,j,r)] for r in range(1,WPR_dim+1))) >= P_Up[(t,j)] - mdl.sum(del_S * contri_reg_percent * P_max_BESS[s-1] for s in range(1,BESS_dim+1)) - mdl.sum(del_S * contri_reg_percent * P_DA_Energy[(t,j,r)] for r in range(1,WPR_dim+1)) for t in range(1,time_dim+1) for j in range(1,min_dim+1))  # Equation(53)

        mdl.add_constraints((-Robust_percent) * (mdl.sum(del_S * contri_reg_percent * P_max_BESS[s-1] for s in range(1,BESS_dim+1)) + mdl.sum(del_S * contri_reg_percent * P_DA_Energy[(t,j,r)] for r in range(1,WPR_dim+1))) <= P_Down[(t,j)] - mdl.sum(del_S * contri_reg_percent * P_max_BESS[s-1] for s in range(1,BESS_dim+1)) - mdl.sum(del_S * contri_reg_percent * P_DA_Energy[(t,j,r)] for r in range(1,WPR_dim+1)) for t in range(1,time_dim+1) for j in range(1,min_dim+1))  # Equation(54)

        mdl.add_constraints((Robust_percent) * (mdl.sum(del_S * contri_reg_percent * P_max_BESS[s-1] for s in range(1,BESS_dim+1)) + mdl.sum(del_S * contri_reg_percent * P_DA_Energy[(t,j,r)] for r in range(1,WPR_dim+1))) >= P_Down[(t,j)] - mdl.sum(del_S * contri_reg_percent * P_max_BESS[s-1] for s in range(1,BESS_dim+1)) - mdl.sum(del_S * contri_reg_percent * P_DA_Energy[(t,j,r)] for r in range(1,WPR_dim+1)) for t in range(1,time_dim+1) for j in range(1,min_dim+1))   # Equation(54)
        
        mdl.add_constraints((-Robust_percent) * Expected_P_RT_WPR.Cells(t+1,j+1).Value * D_WPR[(t,j,r)] <= P_RT_WPR[(t,j,r)] - Expected_P_RT_WPR.Cells(t+1,j+1).Value * D_WPR[(t,j,r)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for r in range(1,WPR_dim+1))  # Equation(55) / 변동구간 +-10%

        mdl.add_constraints(P_RT_WPR[(t,j,r)] - Expected_P_RT_WPR.Cells(t+1,j+1).Value * D_WPR[(t,j,r)] <= (Robust_percent) * Expected_P_RT_WPR.Cells(t+1,j+1).Value * D_WPR[(t,j,r)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for r in range(1,WPR_dim+1))  # Equation(55) / 변동구간 +-10%
    
    else:    
        mdl.add_constraints(P_RT_WPR[(t,j,r)] <= Expected_P_RT_WPR.Cells(t+1,j+1).Value for t in range(1,time_dim+1) for j in range(1,min_dim+1) for r in range(1,WPR_dim+1))  # Equation(55) / 변동구간 +-10%

    return mdl

    
### Main Program    
if __name__ == '__main__':
    for r in Robust_percents:
        for cont in contri_reg_percents:
            mdl = build_optimization_model(Robust_percent=r, contri_reg_percent=cont) # Optimization model
            mdl.print_information() # Model information output
            s = mdl.solve(log_output=True) # Solve optimization model
    
            if s: # With solution
                obj = mdl.objective_value
                mdl.get_solve_details()
                print("* Total cost=%g" % obj)
                print("*Gap tolerance = ", mdl.parameters.mip.tolerances.mipgap.get())
                
                data = [v.name.split('_') + [s.get_value(v)] for v in mdl.iter_variables()] # Save variable result
                frame = pd.DataFrame(data, columns=['var', 'index1', 'index2', 'index3', 'value']) # Save index 2 of variable result with only time
                frame.to_excel(os.getcwd()+"\\Result\\variable_result_Robust_"+str(r)+"_"+"Contri_"+str(cont)+".xlsx")       
                
                # Save the CPLEX solution as "solution.json" program output
                with get_environment().get_output_stream(os.getcwd()+"\\Result\\Json_Files\\"+"solution_Robust_"+str(r)+"_"+"Contri_"+str(cont)+".json") as fp: #Save solution with json
                    mdl.solution.export(fp, "json")
                print("model has solution")
        
            else: # With no solution
                print("* model has no solution")
                 
    print("Program is terminated")
    excel.Quit()
    
import json
import win32com.client as win32
import pandas as pd
import numpy as np
import os
import matplotlib.pyplot as plt

############## Input
Robust_percents = [0, 0.2, 0.4]   # Robust percent (0~1) / (Delta R, Variation interval in robust optimization)
contri_reg_percents = [0, 0.2, 0.4, 0.6, 0.8, 1] # Gamma, Maximum Contribution ratio of providing deployed up and down power in the real-time operation

### Write Result
def result_optimization_model(obj, DataFrame):
    
    frame = DataFrame
    excel = win32.Dispatch("Excel.Application")
    
    wb1 = excel.Workbooks.Open(os.getcwd()+"\\Data\\robust model_data.xlsx")
    Price_DA = wb1.Sheets("Price_DA")              # Day-ahead prices
    Price_DARS = wb1.Sheets("Price_DARS")          # Day-ahead regulation prices (Reserve prices)
    Price_RT = wb1.Sheets("Price_RT")              # Real time prices (Up regulation prices)
    Price_RTRS = wb1.Sheets("Price_RTRS")          # Real time regulation prices (Down regulation prices)
    Expected_P_RT_WPR = wb1.Sheets("Expected_P_RT_WPR")  # Expected generation of renewable energy in the real-time operation

    ### 파라미터 설정
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
    
    wb_result = excel.Workbooks.Open(wb_result_file)
    ws1 = wb_result.Worksheets("Optimization Result")
    ws2 = wb_result.Worksheets("Day-Ahead")
    
    ### Sheet 1  
    # Total Revenue
    ws1.Cells(1,2).Value = "Optimization Result"
    ws1.Cells(2,1).Value = "Total Revenue [$]"
    ws1.Cells(2,2).Value = float(obj)
       
    # AV_RO
    ws1.Cells(3,1).Value = "Income in real-time [$]"
    ws1.Cells(3,2).Value = frame.loc[frame['var']=="AV-RO"]['index1'].sum()
    
    ### Case Result
    ws1.Cells(7,1).Value = "Contribution"
    
    ## Day-ahead Result
    ws1.Cells(5,2).Value = "Day-ahead"
    
    #Result for BESS
    for s in range(1, BESS_dim+1):
        if s == 1:
            BESS_DA_DCH_cost_result = []
            BESS_DA_CH_cost_result = []
            BESS_DA_DCH_price_result = []
            BESS_DA_CH_price_result = []
            BESS_DA_RS_price_result=[]
            
        BESS_DA_DCH_cost_result.append(0)
        BESS_DA_CH_cost_result.append(0)
        BESS_DA_DCH_price_result.append(0)
        BESS_DA_CH_price_result.append(0)
        BESS_DA_RS_price_result.append(0)
        for t in range(1, time_dim+1):
            for j in range(1, min_dim+1):
                BESS_DA_DCH_cost_result[s-1] += del_S * Marginal_cost_DCH[s-1] * frame.loc[frame['var']=="P-DA-DCH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum()
                BESS_DA_CH_cost_result[s-1] += del_S * Marginal_cost_CH[s-1] * frame.loc[frame['var']=="P-DA-CH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum()
                    
                BESS_DA_DCH_price_result[s-1] += del_S * Price_DA.Cells(t+1,2).Value * frame.loc[frame['var']=="P-DA-DCH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum()
                BESS_DA_CH_price_result[s-1] += del_S * Price_DA.Cells(t+1,2).Value * frame.loc[frame['var']=="P-DA-CH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum()
                
                BESS_DA_RS_price_result[s-1] += del_S * Price_DARS.Cells(t+1,2).Value * (frame.loc[frame['var']=="P-RS-CH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum()
                                                                                       +frame.loc[frame['var']=="P-RS-DCH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum())
                    
        ws1.Cells(6,2+s-1).Value = "BESS#"+str(s)
        ws1.Cells(7,2+s-1).Value = BESS_DA_RS_price_result[s-1]+(BESS_DA_DCH_price_result[s-1]-BESS_DA_CH_price_result[s-1])-(BESS_DA_DCH_cost_result[s-1] + BESS_DA_CH_cost_result[s-1])
        
    #Result for Wind
    for r in range(1, WPR_dim+1):
        if r == 1:
            Wind_DA_cost_result = []
            Wind_DA_price_result = []
            Wind_DA_RS_price_result=[]
        Wind_DA_cost_result.append(0)
        Wind_DA_price_result.append(0)
        Wind_DA_RS_price_result.append(0)
        for t in range(1, time_dim+1):
            for j in range(1, min_dim+1):
                Wind_DA_cost_result[r-1] += del_S * Marginal_cost_WPR[r-1] * frame.loc[frame['var']=="P-DA-Energy"]['value'][min_dim*WPR_dim*(t-1)+WPR_dim*(j-1)+r-1:min_dim*WPR_dim*(t-1)+WPR_dim*(j-1)+r].sum()
                Wind_DA_price_result[r-1] += del_S * Price_DA.Cells(t+1,2).Value * frame.loc[frame['var']=="P-DA-Energy"]['value'][min_dim*WPR_dim*(t-1)+WPR_dim*(j-1)+r-1:min_dim*WPR_dim*(t-1)+WPR_dim*(j-1)+r].sum()
                Wind_DA_RS_price_result[r-1] += del_S * Price_DARS.Cells(t+1,2).Value * frame.loc[frame['var']=="P-RS-RES"]['value'][min_dim*WPR_dim*(t-1)+WPR_dim*(j-1)+r-1:min_dim*WPR_dim*(t-1)+WPR_dim*(j-1)+r].sum()
                    
        ws1.Cells(6,2+BESS_dim+r-1).Value = "Wind#"+str(robust)
        ws1.Cells(7,2+BESS_dim+r-1).Value = Wind_DA_RS_price_result[r-1]+Wind_DA_price_result[r-1]-Wind_DA_cost_result[r-1]
    
    ## Real-time Result
    ws1.Cells(5,2+BESS_dim+WPR_dim).Value = "Real-time"
    #Result for BESS
    for s in range(1, BESS_dim+1):
        if s == 1:
            BESS_RT_DCH_cost_result = []
            BESS_RT_CH_cost_result = []
            BESS_RT_DCH_price_result = []
            BESS_RT_CH_price_result = []
        BESS_RT_DCH_cost_result.append(0)
        BESS_RT_CH_cost_result.append(0)
        BESS_RT_DCH_price_result.append(0)
        BESS_RT_CH_price_result.append(0)
        for t in range(1, time_dim+1):
            for j in range(1, min_dim+1):
                
                BESS_RT_DCH_cost_result[s-1] += del_S * Marginal_cost_DCH[s-1] * (frame.loc[frame['var']=="P-Up-DCH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum()
                                                                          + frame.loc[frame['var']=="P-Down-DCH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum()) 
                BESS_RT_CH_cost_result[s-1] += del_S * Marginal_cost_CH[s-1] * (frame.loc[frame['var']=="P-Up-CH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum()
                                                                        +frame.loc[frame['var']=="P-Down-CH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum())
                    
                BESS_RT_DCH_price_result[s-1] += (Price_RT.Cells(t+1,j+1).Value * del_S * frame.loc[frame['var']=="P-Up-DCH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum()
                                                    + Price_RT.Cells(t+1,j+1).Value * del_S * frame.loc[frame['var']=="P-Down-DCH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum()
                                                    - Price_RTRS.Cells(t+1,j+1).Value * del_S * frame.loc[frame['var']=="P-Down-DCH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum())
                BESS_RT_CH_price_result[s-1] += (Price_RT.Cells(t+1,j+1).Value * del_S * frame.loc[frame['var']=="P-Up-CH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum()
                                                    + Price_RT.Cells(t+1,j+1).Value * del_S * frame.loc[frame['var']=="P-Down-CH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum()
                                                    - Price_RTRS.Cells(t+1,j+1).Value * del_S * frame.loc[frame['var']=="P-Down-CH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum())
                    
        ws1.Cells(6,2+BESS_dim+WPR_dim+s-1).Value = "BESS#"+str(s)
        ws1.Cells(7,2+BESS_dim+WPR_dim+s-1).Value = (BESS_RT_DCH_price_result[s-1]-BESS_RT_CH_price_result[s-1])-(BESS_RT_DCH_cost_result[s-1] + BESS_RT_CH_cost_result[s-1])
        
    #Result for Wind
    for r in range(1, WPR_dim+1):
        if r == 1:
            Wind_RT_cost_result = []
            Wind_RT_price_result = []
        Wind_RT_cost_result.append(0)
        Wind_RT_price_result.append(0)
        for t in range(1, time_dim+1):
            for j in range(1, min_dim+1):
                Wind_RT_cost_result[r-1] += del_S * Marginal_cost_WPR[r-1] * frame.loc[frame['var']=="P-Up-RES"]['value'][min_dim*WPR_dim*(t-1)+WPR_dim*(j-1)+r-1:min_dim*WPR_dim*(t-1)+WPR_dim*(j-1)+r].sum()
                Wind_RT_price_result[r-1] += (Price_RT.Cells(t+1,j+1).Value * del_S * frame.loc[frame['var']=="P-Up-RES"]['value'][min_dim*WPR_dim*(t-1)+WPR_dim*(j-1)+r-1:min_dim*WPR_dim*(t-1)+WPR_dim*(j-1)+r].sum()
                                              -Price_RT.Cells(t+1,j+1).Value * del_S * frame.loc[frame['var']=="P-Down-RES"]['value'][min_dim*WPR_dim*(t-1)+WPR_dim*(j-1)+r-1:min_dim*WPR_dim*(t-1)+WPR_dim*(j-1)+r].sum()
                                              -Price_RT.Cells(t+1,j+1).Value * del_S * frame.loc[frame['var']=="P-RES-imb"]['value'][min_dim*WPR_dim*(t-1)+WPR_dim*(j-1)+r-1:min_dim*WPR_dim*(t-1)+WPR_dim*(j-1)+r].sum()
                                              +Price_RTRS.Cells(t+1,j+1).Value * del_S * frame.loc[frame['var']=="P-Down-RES"]['value'][min_dim*WPR_dim*(t-1)+WPR_dim*(j-1)+r-1:min_dim*WPR_dim*(t-1)+WPR_dim*(j-1)+r].sum()
                                              )
                    
        ws1.Cells(6,2+BESS_dim+WPR_dim+BESS_dim+r-1).Value = "Wind#"+str(robust)
        ws1.Cells(7,2+BESS_dim+WPR_dim+BESS_dim+r-1).Value = Wind_RT_price_result[r-1]-Wind_RT_cost_result[r-1]    
    
    print("Optimization Result Calculation Done!")   
    
    wb_result.Save()
    excel.Quit()
    
    ### Graph for Energy of BESS
    E_BESS_result1 = [0] * 24
    E_BESS_result2 = [0] * 24     
    for s in range(1, BESS_dim+1):               
        for t in range(1, time_dim+1):
            for j in range(1, min_dim+1):    
                
                if s == 1:
                    E_BESS_result1[t-1] += frame.loc[frame['var']=="E-BESS"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum()

                elif s == 2:   
                    E_BESS_result2[t-1] += frame.loc[frame['var']=="E-BESS"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum()
    
    for i in range(len(E_BESS_result1)):
        E_BESS_result1[i] /= 12
        E_BESS_result2[i] /= 12    
    
    x = np.arange(0, time_dim, 1)      
    plt.plot(x, E_BESS_result1, 'g', linewidth = 5, label ='E_BESS#1')
    plt.plot(x, E_BESS_result2, 'darkviolet', linewidth = 5, label ='E_BESS#2')
    plt.xlabel('Time (h)', fontweight= 'bold', size = 14)
    plt.ylabel('Energy (MWh)', fontweight= 'bold', size = 14)
    plt.legend()
    plt.title('Energy of BESS ($\Delta R$ : ' + str(robust) + ' ' + '/' + ' ' + '$\gamma$ : ' + str(cont) + ')')
    plt.xticks(np.arange(0, time_dim, 1))  
    plt.yticks(np.arange(0 , 30, 5))
    plt.show(block=False)
    plt.savefig(figure_folder+"Energy of BESS(Robust_"+str(robust)+"_"+"Contri_"+str(cont)+").png")
    plt.pause(1)
    plt.close()
 
    ### Graph for Day-ahead Power
    BESS1_P_DA_DCH_result = [0] * 24
    BESS1_P_DA_CH_result = [0] * 24
    BESS2_P_DA_DCH_result = [0] * 24
    BESS2_P_DA_CH_result = [0] * 24
    P_DA_WPR_result = [0] * 24     
    
    for s in range(1, BESS_dim+1):               
        for t in range(1, time_dim+1):
            for j in range(1, min_dim+1):    
                
                if s == 1:
                    BESS1_P_DA_DCH_result[t-1] += frame.loc[frame['var']=="P-DA-DCH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum()
                    BESS1_P_DA_CH_result[t-1] -= frame.loc[frame['var']=="P-DA-CH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum()
                        
                elif s == 2:
                    BESS2_P_DA_DCH_result[t-1] += frame.loc[frame['var']=="P-DA-DCH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum()
                    BESS2_P_DA_CH_result[t-1] -= frame.loc[frame['var']=="P-DA-CH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum()
    
    for r in range(1, WPR_dim+1):               
            for t in range(1, time_dim+1):
                for j in range(1, min_dim+1):    
                    
                    P_DA_WPR_result[t-1] += frame.loc[frame['var']=="P-DA-Energy"]['value'][min_dim*WPR_dim*(t-1)+WPR_dim*(j-1)+r-1:min_dim*WPR_dim*(t-1)+WPR_dim*(j-1)+r].sum() 
    
    for i in range(len(BESS1_P_DA_DCH_result)):
        BESS1_P_DA_DCH_result[i] /= 12
        BESS1_P_DA_CH_result[i] /= 12    
        BESS2_P_DA_DCH_result[i] /= 12
        BESS2_P_DA_CH_result[i] /= 12 
        P_DA_WPR_result[i] /= 12             
                
    x = np.arange(0, time_dim, 1)
    bottom_P_DA_RES = np.add(BESS1_P_DA_DCH_result, BESS2_P_DA_DCH_result)
    plt.bar(x, BESS1_P_DA_DCH_result, color = 'blue', linewidth = 5, label ='P_BESS#1_DA')
    plt.bar(x, BESS1_P_DA_CH_result, color = 'blue', linewidth = 5)             
    plt.bar(x, BESS2_P_DA_DCH_result, color = 'red', linewidth = 5, label ='P_BESS#2_DA', bottom = BESS1_P_DA_DCH_result)
    plt.bar(x, BESS2_P_DA_CH_result, color = 'red', linewidth = 5, bottom = BESS1_P_DA_CH_result)
    plt.bar(x, P_DA_WPR_result, color = 'lime', linewidth = 5, label ='P_Wind_DA', bottom = bottom_P_DA_RES)            
    plt.xlabel('Time (h)', fontweight= 'bold', size = 14)
    plt.ylabel('Power (MW)', fontweight= 'bold', size = 14)
    plt.legend()
    plt.title('Day-ahead Power ($\Delta R$ : ' + str(robust) + ' ' + '/' + ' ' + '$\gamma$ : ' + str(cont) + ')')
    plt.xticks(np.arange(0, time_dim, 1))  
    plt.yticks(np.arange(-10, 11, 2))
    #plt.grid(True, axis='y')
    plt.axhline(0, color='black', linewidth=1)
    plt.show(block=False)
    plt.savefig(figure_folder+"Day-ahead Power(Robust_"+str(robust)+"_"+"Contri_"+str(cont)+").png")
    plt.pause(1)
    plt.close() 

    ### Graph for Day-ahead Reserve Power
    BESS1_P_RS_DCH_result = [0] * 24
    BESS1_P_RS_CH_result = [0] * 24
    BESS2_P_RS_DCH_result = [0] * 24
    BESS2_P_RS_CH_result = [0] * 24
    P_RS_WPR_result = [0] * 24     
    
    for s in range(1, BESS_dim+1):               
        for t in range(1, time_dim+1):
            for j in range(1, min_dim+1):    
                
                if s == 1:
                    BESS1_P_RS_DCH_result[t-1] += frame.loc[frame['var']=="P-RS-DCH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum()
                    BESS1_P_RS_CH_result[t-1] += frame.loc[frame['var']=="P-RS-CH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum()
                        
                elif s == 2:
                    BESS2_P_RS_DCH_result[t-1] += frame.loc[frame['var']=="P-RS-DCH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum()
                    BESS2_P_RS_CH_result[t-1] += frame.loc[frame['var']=="P-RS-CH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum()
    
    for r in range(1, WPR_dim+1):               
            for t in range(1, time_dim+1):
                for j in range(1, min_dim+1):    
                    
                    P_RS_WPR_result[t-1] += frame.loc[frame['var']=="P-RS-RES"]['value'][min_dim*WPR_dim*(t-1)+WPR_dim*(j-1)+r-1:min_dim*WPR_dim*(t-1)+WPR_dim*(j-1)+r].sum() 
    
    max_RS = []
    
    for i in range(len(BESS1_P_RS_DCH_result)):
        BESS1_P_RS_DCH_result[i] /= 12
        BESS1_P_RS_CH_result[i] /= 12    
        BESS2_P_RS_DCH_result[i] /= 12
        BESS2_P_RS_CH_result[i] /= 12 
        P_RS_WPR_result[i] /= 12             
        max_RS.append(BESS1_P_RS_DCH_result[i]+BESS1_P_RS_CH_result[i]+BESS2_P_RS_DCH_result[i]+BESS2_P_RS_CH_result[i]+P_RS_WPR_result[i])
    
    x = np.arange(0, time_dim, 1)
    plt.bar(x, BESS1_P_RS_DCH_result, color = 'blue', linewidth = 5, label ='P_BESS#1_RS_DCH')
    plt.bar(x, BESS1_P_RS_CH_result, color = 'lightsteelblue', linewidth = 5, label ='P_BESS#1_RS_CH', bottom = BESS1_P_RS_DCH_result)
    plt.bar(x, BESS2_P_RS_DCH_result, color = 'red', linewidth = 5, label ='P_BESS#2_RS_DCH', bottom = [sum(x) for x in zip(BESS1_P_RS_DCH_result, BESS1_P_RS_CH_result)])
    plt.bar(x, BESS2_P_RS_CH_result, color = 'lightcoral', linewidth = 5, label ='P_BESS#2_RS_CH', bottom = [sum(x) for x in zip(BESS1_P_RS_DCH_result, BESS1_P_RS_CH_result, BESS2_P_RS_DCH_result)])
    plt.bar(x, P_RS_WPR_result, color = 'lime', linewidth = 5, label ='P_Wind_RS', bottom = [sum(x) for x in zip(BESS1_P_RS_DCH_result, BESS1_P_RS_CH_result, BESS2_P_RS_DCH_result, BESS2_P_RS_CH_result)])
    plt.xlabel('Time (h)', fontweight= 'bold', size = 14)
    plt.ylabel('Power (MW)', fontweight= 'bold', size = 14)
    plt.legend()
    plt.title('Day-ahead Reserve ($\Delta R$ : ' + str(robust) + ' ' + '/' + ' ' + '$\gamma$ : ' + str(cont) + ')')
    plt.xticks(np.arange(0, time_dim, 1))  
    plt.yticks(np.arange(0, round(max(max_RS))+2, 1))
    #plt.grid(True, axis='y')
    plt.axhline(0, color='black', linewidth=1)
    plt.show(block=False)
    plt.savefig(figure_folder+"Day-ahead Reserve(Robust_"+str(robust)+"_"+"Contri_"+str(cont)+").png")
    plt.pause(1)
    plt.close()  

    ### Graph for Up-regulation Power
    BESS1_P_Up_DCH_result = [0] * 24
    BESS1_P_Up_CH_result = [0] * 24
    BESS2_P_Up_DCH_result = [0] * 24
    BESS2_P_Up_CH_result = [0] * 24
    P_Up_RES_result = [0] * 24     
    
    for s in range(1, BESS_dim+1):               
        for t in range(1, time_dim+1):
            for j in range(1, min_dim+1):    
                
                if s == 1:
                    BESS1_P_Up_DCH_result[t-1] += frame.loc[frame['var']=="P-Up-DCH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum()
                    BESS1_P_Up_CH_result[t-1] -= frame.loc[frame['var']=="P-Up-CH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum()
                        
                elif s == 2:
                    BESS2_P_Up_DCH_result[t-1] += frame.loc[frame['var']=="P-Up-DCH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum()
                    BESS2_P_Up_CH_result[t-1] -= frame.loc[frame['var']=="P-Up-CH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum()
    
    for r in range(1, WPR_dim+1):               
            for t in range(1, time_dim+1):
                for j in range(1, min_dim+1):    
                    
                    P_Up_RES_result[t-1] += frame.loc[frame['var']=="P-Up-RES"]['value'][min_dim*WPR_dim*(t-1)+WPR_dim*(j-1)+r-1:min_dim*WPR_dim*(t-1)+WPR_dim*(j-1)+r].sum() 
    
    for i in range(len(BESS1_P_Up_DCH_result)):
        BESS1_P_Up_DCH_result[i] /= 12
        BESS1_P_Up_CH_result[i] /= 12    
        BESS2_P_Up_DCH_result[i] /= 12
        BESS2_P_Up_CH_result[i] /= 12 
        P_Up_RES_result[i] /= 12             
                
    x = np.arange(0, time_dim, 1)
    bottom_P_Up_RES = np.add(BESS1_P_Up_DCH_result, BESS2_P_Up_DCH_result)
    plt.bar(x, BESS1_P_Up_DCH_result, color = 'blue', linewidth = 5, label ='P_BESS#1_Up')
    plt.bar(x, BESS1_P_Up_CH_result, color = 'blue', linewidth = 5)             
    plt.bar(x, BESS2_P_Up_DCH_result, color = 'red', linewidth = 5, label ='P_BESS#2_Up', bottom = BESS1_P_Up_DCH_result)
    plt.bar(x, BESS2_P_Up_CH_result, color = 'red', linewidth = 5, bottom = BESS1_P_Up_CH_result)
    plt.bar(x, P_Up_RES_result, color = 'lime', linewidth = 5, label ='P_Wind_Up', bottom = bottom_P_Up_RES)            
    plt.xlabel('Time (h)', fontweight= 'bold', size = 14)
    plt.ylabel('Power (MW)', fontweight= 'bold', size = 14)
    plt.legend()
    plt.title('Real-time deployed Up Power ($\Delta R$ : ' + str(robust) + ' ' + '/' + ' ' + '$\gamma$ : ' + str(cont) + ')')
    plt.xticks(np.arange(0, time_dim, 1))  
    plt.yticks(np.arange(-10, 11, 2))
    #plt.grid(True, axis='y')
    plt.axhline(0, color='black', linewidth=1)
    plt.show(block=False)
    plt.savefig(figure_folder+"Real-time deployed Up Power(Robust_"+str(robust)+"_"+"Contri_"+str(cont)+").png")
    plt.pause(1)
    plt.close()  

    ### Graph for Down-regulation Power
    BESS1_P_Down_DCH_result = [0] * 24
    BESS1_P_Down_CH_result = [0] * 24
    BESS2_P_Down_DCH_result = [0] * 24
    BESS2_P_Down_CH_result = [0] * 24
    P_Down_RES_result = [0] * 24     
    
    for s in range(1, BESS_dim+1):               
        for t in range(1, time_dim+1):
            for j in range(1, min_dim+1):    
                
                if s == 1:
                    BESS1_P_Down_DCH_result[t-1] += frame.loc[frame['var']=="P-Down-DCH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum()
                    BESS1_P_Down_CH_result[t-1] -= frame.loc[frame['var']=="P-Down-CH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum()
                        
                elif s == 2:
                    BESS2_P_Down_DCH_result[t-1] += frame.loc[frame['var']=="P-Down-DCH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum()
                    BESS2_P_Down_CH_result[t-1] -= frame.loc[frame['var']=="P-Down-CH"]['value'][min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s-1:min_dim*BESS_dim*(t-1)+BESS_dim*(j-1)+s].sum()
    
    for r in range(1, WPR_dim+1):               
            for t in range(1, time_dim+1):
                for j in range(1, min_dim+1):    
                    
                    P_Down_RES_result[t-1] += frame.loc[frame['var']=="P-Down-RES"]['value'][min_dim*WPR_dim*(t-1)+WPR_dim*(j-1)+r-1:min_dim*WPR_dim*(t-1)+WPR_dim*(j-1)+r].sum() 
    
    for i in range(len(BESS1_P_Down_DCH_result)):
        BESS1_P_Down_DCH_result[i] /= 12
        BESS1_P_Down_CH_result[i] /= 12    
        BESS2_P_Down_DCH_result[i] /= 12
        BESS2_P_Down_CH_result[i] /= 12 
        P_Down_RES_result[i] /= 12             
                
    x = np.arange(0, time_dim, 1)
    bottom_P_Down_RES = np.add(BESS1_P_Down_DCH_result, BESS2_P_Down_DCH_result)
    plt.bar(x, BESS1_P_Down_DCH_result, color = 'blue', linewidth = 5, label ='P_BESS#1_Down')
    plt.bar(x, BESS1_P_Down_CH_result, color = 'blue', linewidth = 5)             
    plt.bar(x, BESS2_P_Down_DCH_result, color = 'red', linewidth = 5, label ='P_BESS#2_Down', bottom = BESS1_P_Down_DCH_result)
    plt.bar(x, BESS2_P_Down_CH_result, color = 'red', linewidth = 5, bottom = BESS1_P_Down_CH_result)
    plt.bar(x, P_Down_RES_result, color = 'lime', linewidth = 5, label ='P_Wind_Down', bottom = bottom_P_Down_RES)            
    plt.xlabel('Time (h)', fontweight= 'bold', size = 14)
    plt.ylabel('Power (MW)', fontweight= 'bold', size = 14)
    plt.legend()
    plt.title('Real-time deployed Down Power ($\Delta R$ : ' + str(robust) + ' ' + '/' + ' ' + '$\gamma$ : ' + str(cont) + ')')
    plt.xticks(np.arange(0, time_dim, 1))  
    plt.yticks(np.arange(-10, 11, 2))
    #plt.grid(True, axis='y')
    plt.axhline(0, color='black', linewidth=1)
    plt.show(block=False)
    plt.savefig(figure_folder+"Real-time deployed Down Power(Robust_"+str(robust)+"_"+"Contri_"+str(cont)+").png")
    plt.pause(1)
    plt.close()   
#####################

def writing_results(json_file, result_file):
    with open(json_file,'r') as f:
        json_data = json.load(f)
    obj = json_data['CPLEXSolution']['header']['objectiveValue']
    frame = pd.read_excel(result_file)
    result_optimization_model(obj,frame)
    print("Results Done")

if __name__ == "__main__":
    for robust in Robust_percents:
        for cont in contri_reg_percents:
            json_file = os.getcwd()+"\\Result\\Json_Files\\"+"solution_Robust_"+str(robust)+"_"+"Contri_"+str(cont)+".json"
            result_file= os.getcwd()+"\\Result\\variable_result_Robust_"+str(robust)+"_"+"Contri_"+str(cont)+".xlsx"
            wb_result_file =  os.getcwd()+"\\Result\\robust model_result_Robust_"+str(robust)+"_"+"Contri_"+str(cont)+".xlsx"
            figure_folder = os.getcwd()+"\\Figure\\Robust_"+str(robust)+"_"+"Contri_"+str(cont)+"\\"
            writing_results(json_file, result_file)
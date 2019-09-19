import xlrd
from pulp import * 
import scipy
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import numpy as np

df = pd.read_excel('Copy of Project_sitestrategy_finance.v2.xlsx', sheet_name='Sheet2')
print(df)
problem_name = 'optimize FTEs'
ftes=1000

def optimize_Fte(ftes):
    #create the LP object, set up as a minimization problem --> since we want to minimize the costs
    prob = pulp.LpProblem(problem_name, pulp.LpMinimize)

#create decision variables
    decision_variables = []
    for rownum, row in df.iterrows():
        variable = str('x' + str(rownum))
        variable = pulp.LpVariable(str(variable), lowBound = 0, cat= 'Integer')
        decision_variables.append(variable)
        #print ("Total number of decision_variables: " + str(len(decision_variables)))
        #print(decision_variables)
        
#create objective Function -minimize the costs for the move
    total_cost = ""
    for rownum, row in df.iterrows():
        for i, j in enumerate(decision_variables):
            if rownum == i:
                total_cost= 2.42*(row['rent']*j) + total_cost
    #print(total_cost)
            

    prob += total_cost
        #print ("Optimization function: " + str(total_cost))	
 
#create constrains -
        
    prob+= decision_variables[0]+decision_variables[1]+decision_variables[2]+decision_variables[3]+decision_variables[4]+decision_variables[5]+decision_variables[6]+decision_variables[7]==ftes
    prob+= decision_variables[2] <= 100
    prob+= decision_variables[1] <= decision_variables[0]
    prob+= decision_variables[5] <= decision_variables[0]
    prob+= decision_variables[3] <= 29
    prob+= decision_variables[4] <= 45
    prob+= decision_variables[6] <= decision_variables[0]
    prob+= decision_variables[7] <= decision_variables[0]
    

#now run optimization
    prob.writeLP(problem_name + ".lp" )
    optimization_result = prob.solve()
    assert optimization_result == pulp.LpStatusOptimal
    print("Status:", LpStatus[prob.status])
    print("Optimal Solution to the problem: ", value(prob.objective)*12*5)
    print ("Individual decision_variables: ")
    s=0
    for v in prob.variables():
        s=s+v.varValue
        print(v.name, "=", v.varValue)
    print(s)
if __name__ == "__main__":
    optimize_Fte(ftes)

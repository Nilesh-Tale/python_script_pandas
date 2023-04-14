import pandas as pd

def teamwise():
    
    #opening file
    filename = "E:\\Coding\\Python\\Agile RecruiTech LLP_ Assignment\\InputData.xlsx"
    pd.ExcelFile(filename).sheet_names

    #converting input file sheet1 into pandas dataframe
    df1 = pd.read_excel(filename, sheet_name = "Input_sheet1")
    df1.drop(['S No'], axis =1, inplace = True)
    df1.rename(columns = {'User ID': 'uid', 'Team Name': 'team_name'}, inplace = True)

    #converting input file sheet2 into pandas dataframe
    df2 = pd.read_excel(filename, sheet_name = "Input_sheet2")
    df2.drop(['S No'], axis =1, inplace = True)
    
    #merging both the dataframes
    df = pd.merge(df1, df2, on='uid')
    df.drop(['Name'], axis = 1, inplace = True)

    #creating unique list of names
    team_arr = df.team_name.unique()
    team_list = team_arr.tolist()
    
    #Performing operations on merged dataframe
    #creating lists
    avg_stat = []
    avg_reas = []
    sum_s_r=[]

    #generation of unique names of teams
    team_arr = df.team_name.unique()
    team_list = team_arr.tolist()

    #creating new dataframe for desired result
    df3 = pd.DataFrame(team_list, columns = ['Thinking Teams Leaderboard'])
 
    
    for team in team_list:
        select_team = df.loc[df['team_name'] == team]
        stat = round(select_team.loc[:, 'total_statements'].mean(), 2)
        reas = round(select_team.loc[:, 'total_reasons'].mean(),2)
        sum_sr = stat + reas
    
        sum_s_r.append(sum_sr)
        avg_stat.append(stat)
        avg_reas.append(reas)
    
    #adding new columns to result dataframe
    df3.insert(1, "Average Statements", avg_stat)
    df3.insert(2, "Average Reasons", avg_reas)

    #creating extra column that contains sum of average statements and average reasons
    df3.insert(3, "Sum of avg. statements and Reasons", sum_s_r)
    
    #sorting resultant dataframe according to extra column that we generated
    df4 = df3.sort_values(by = "Sum of avg. statements and Reasons", ascending = False)

    #creating rank column
    rank_no = []
    for i in range(1, len(df4.index)+1):
        rank_no.append(i)
    
    #inserting rank column at the start
    df4.insert(0, "Team Rank", rank_no)

    #reseting pandas dataframe index
    df4.reset_index(inplace = True, drop = True)

    #dropping the extra column that we generated
    df4.drop(['Sum of avg. statements and Reasons'], axis = 1, inplace =True)

    #returning the desired output dataframe
    return df4


def individual():
    filename="E:\\Coding\\Python\\Agile RecruiTech LLP_ Assignment\\InputData.xlsx"

    #reading the required input sheet from given file
    df = pd.read_excel(filename, sheet_name = "Input_sheet2")
    df.drop(['S No'], axis =1, inplace = True)
    df.rename(columns = {'User ID': 'uid'}, inplace = True)
    
    #creating extra column that contains the addition of values from total_statements and total_reasons columns
    df['Sum'] = df['total_statements'] + df['total_reasons']

    #sorting the dataframe in descending order with new generated sum column
    df1 = df.sort_values(by = "Sum", ascending = False)

    #creating rank_no list and appending in resultant dataframe
    rank_no = []
    for i in range(1, len(df1.index)+1):
        rank_no.append(i)
    df1.insert(0, "Rank", rank_no)

    #reseting index
    df1.reset_index(inplace = True, drop = True)

    #dropping extra generated column
    df1.drop(['Sum'], axis = 1, inplace =True)
    
    #returning the desired resultant dataframe
    return df1

df1=teamwise()
df2=individual()

#writing the obtained dataframes to excel file into two different sheets
with pd.ExcelWriter("E:\\Coding\\Python\\Agile RecruiTech LLP_ Assignment\\output.xlsx") as writer:
    df1.to_excel(writer, sheet_name="Leaderboard Individual (Output)", index=False)
    df2.to_excel(writer, sheet_name="Leaderboard TeamWise (Output)", index=False)
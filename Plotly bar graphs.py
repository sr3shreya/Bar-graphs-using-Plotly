from plotly.subplots import make_subplots
import plotly.graph_objects as go
import pandas as pd
import plotly
import time
import os
from datetime import date

desktop = os.path.join(os.path.join(os.environ['USERPROFILE']),'Desktop','SCJ','Agent_data_graph')
t_date = str(date.today())
test_path = desktop + '\\' + t_date
if not os.path.exists(test_path):
    os.makedirs(test_path)
    os.makedirs(test_path + '\\Graph_Images')
    os.makedirs(test_path + '\\Graph_HTML')
    #os.makedirs(test_path+'\\Graph_Images\\%Count')
    #os.makedirs(test_path+'\\Graph_Images\\%Coverage')

def per_count_image(q,figure, temp_var):

    '''figure.update_layout(
        height=1000,
        width=2000,
        showlegend=False

    )'''
    if temp_var=='Server':
        figure.update_layout(
            height=900,
            width=2000,
            showlegend=False
        )
    elif temp_var=='Endpoint':
        figure.update_layout(
            height=600,
            width=1800,
            showlegend=False
        )

    figure.update_layout(title="Change in % Count for " + str(q)+ " "+ temp_var )
    figure.write_image(test_path + '\\Graph_Images\\' + str(q) + ' '+ str(temp_var)+' _%count.png')

def per_coverage_image(j,figure, temp_var):
    '''figure.update_layout(

        height=800,
        width=1500,
        showlegend=False
    )'''
    if temp_var=='Server':
        figure.update_layout(
            height=550,
            width=1400,
            showlegend=False

        )
    elif temp_var=='Endpoint':
        figure.update_layout(
            height=500,
            width=1400,
            showlegend=False
        )

    figure.update_layout(title="Change in % Coverage for " + str(j)+ " "+ temp_var)
    figure.write_image(test_path + '\\Graph_Images\\' + str(j)+' '+ str(temp_var) + ' _%coverage.png')

def per_count(p,q,r,df1):

    df1[p] = df1[p][1:].astype(float).round(decimals=3)
    '''c=[]

    c.append(df[p].nlargest(2))
    #c.append(df[p].nsmallest(1))
    print(c)
    '''
    df_ss = pd.DataFrame()
    c = df1[p].nlargest(1)
    c=c.append(df1[p].nsmallest(1))
    key_s = list(c.keys())
    for j in key_s:
        temp = df1.loc[[j]]
        temp1 = temp[['OS','SCCM',q,r]].rename(
            columns={'SCCM': 'OS Count',q: 'Count Previous Week',r: 'Count This Week'})
        df_ss = df_ss.append(temp1,ignore_index=True)
    '''
    for i in c:
        id=i.index[0]
        print(id)
        df_temp1=df.loc[[id]]
        temp22=df_temp1[['OS','SCCM',q,r]].rename(columns={'SCCM': 'OS Count',q: 'Count Previous Week',r: 'Count This Week'})
        print(temp22)
        df_ss=df_ss.append(temp22, ignore_index=True)
        print(df_ss)
    '''
    colors = ['gold']
    fig = make_subplots(
        rows=2,cols=2,
        shared_xaxes=True,
        specs=[[{"type": "bar"},None],
               [{"type": "table"},None]]
    )
    fig.add_trace(
        go.Table(
            header=dict(values=df_ss.columns.tolist()),
            cells=dict(values=df_ss.to_numpy().T.tolist()),
        ),
        row=2,col=1
    )

    fig.add_trace(
        go.Bar(
            orientation="h",y=df1['OS'],
            x=df1[p],
            text=df1[p],
            textposition='outside',
            marker=dict(color=colors)
        ),
        row=1,col=1
    )
    a=df1.index[2]
    if df1['OS'][a].__contains__('Server') ==True:
        temp_var='Server'
        fig.update_layout(
            height=900,
            width=1900,
            showlegend=False
        )
    else:
        temp_var='Endpoint'
        fig.update_layout(
            height=550,
            width=1650,
            showlegend=False
        )
    per_count_image(q,fig, temp_var)
    '''fig.update_layout(

        height=1000,
        width=2000,
        showlegend=False

    )'''
    '''============================'''
    fig.update_layout(title="Change in % Count for " + str(q)+ " "+ temp_var)
    # fig.show()
    plotly.offline.plot(fig,filename=test_path + '\\Graph_HTML\\' + str(q) + ' '+ str(temp_var)+ ' _%count.html')

def per_coverage(i,j,df2):
    df2[i] = df2[i][1:].astype(float).round(decimals=3)
    colors = ['gold']
    fig = go.Figure(go.Bar(
        orientation="h",y=df2['OS'],
        x=df2[i],
        text=df2[i],
        textposition='outside',
        width= 0.45,
        #marker=dict(color=colors)
        marker = dict(color='rgb(0,128,0)')
    ))
    a = df2.index[2]
    if df2['OS'][a].__contains__('Server') == True:
        temp_var = 'Server'
    else:
        temp_var = 'Endpoint'
    fig.update_layout(title="Change in % Coverage for " + str(j)+ " "+ temp_var)
    plotly.offline.plot(fig,filename=test_path + '\\Graph_HTML\\' + str(j) +' '+ str(temp_var)+ ' _%coverage.html')
    per_coverage_image(j,fig, temp_var)
    # fig.write_image(test_path+'\\'+str(j)+'_%coverage.png' )

ispath = os.path.isdir(test_path)
if ispath:
    df_1 = pd.read_excel(desktop + '\\Agent Data2.xlsx')
    df = df_1.replace(['-'],[0])
    os = df['OS']
    df_endpoint=df.copy()
    df_server=df.copy()
    df_endpoint.drop(df_endpoint[df_endpoint['OS'].str.contains('Server') == True].index,inplace=True)
    df_server.drop(df_server[df_server['OS'].str.contains('Server') == False].index,inplace=True)
    #list_of_dfs=['df_server','df_endpoint']
    percent_c = ['Unnamed: 4','Unnamed: 8','Unnamed: 12','Unnamed: 16']
    percent_coverage = ['Unnamed: 5','Unnamed: 9','Unnamed: 13','Unnamed: 17']
    count_this_week = ['Unnamed: 3','Unnamed: 7','Unnamed: 11','Unnamed: 15']
    agent = ['McAfee','Helix','Rapid7','ForeScout']

#for i in len(a):


#for i in list_of_dfs:

    for p,q,r in zip(percent_c,agent,count_this_week):
        # print(p,q)
        print(p,q,r)
        per_count(p,q,r,df_server)
        per_count(p,q,r,df_endpoint)



        '''df[p] = df[p][1:].astype(float).round(decimals=3)

        df_ss = pd.DataFrame()
        c=df[p].nlargest(2)
        key_s = list(c.keys())
        for j in key_s:
            temp = df.loc[[j]]
            temp1 = temp[['OS','SCCM',q,r]].rename(columns={'SCCM': 'OS Count',q: 'Count Previous Week',r: 'Count This Week'})
            df_ss = df_ss.append(temp1,ignore_index=True)

        colors=['gold']
        fig = make_subplots(
            rows=2, cols=2,
            shared_xaxes=True,
            specs=[[{"type": "bar"}, None],
                   [{"type": "table"}, None]]
        )
        fig.add_trace(
            go.Table(
                header=dict(values=df_ss.columns.tolist()),
                cells=dict(values=df_ss.to_numpy().T.tolist()),
            ),
            row=2, col=1
        )

        fig.add_trace(
           go.Bar(
                    orientation = "h", y = os,
                    x = df[p],
                    text= df[p],
                    textposition='outside',
                    marker=dict(color = colors)
                ),
           row=1, col=1
        )

        fig.update_layout(

            height=1150,
            width=2300,
            showlegend=False

        )

        fig.update_layout(title="Change in % Count " + str(q))
        #fig.show()
        plotly.offline.plot(fig,filename=test_path + '\\' + str(q) + '_%count.html')
        #fig.write_image(test_path+ '\\' + str(q) + '_%count.png')
        '''
    for i,j in zip(percent_coverage,agent):
        print(i,j)
        per_coverage(i,j,df_server)
        per_coverage(i,j,df_endpoint)
        '''df[i]=df[i][1:].astype(float).round(decimals=3)
        fig = go.Figure(go.Bar(
            orientation = "h", y = os,
            x = df[i],
            text= df[i],
            textposition='outside'
        ))


        fig.update_layout(title = "Change in % Coverage for "+ str(j))
        plotly.offline.plot(fig, filename= test_path+'\\'+str(j)+'_%coverage.html')
        #fig.write_image(test_path+'\\'+str(j)+'_%coverage.png' )
        '''
else:
    print("Check the file path")

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import json
import seaborn as sns
import datetime as date
import calendar as cln
from datetime import timedelta
from datetime import datetime, timezone
from dateutil import parser
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import streamlit  as st
from streamlit_option_menu import option_menu
import base64
import csv
import plotly.graph_objs as go
from raceplotly.plots import barplot
from  PIL import Image
from st_aggrid import AgGrid
from streamlit_option_menu import option_menu
import streamlit.components.v1 as html
from collections import deque
                

def get_processing_dfs():

    scope = ['https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
    client = gspread.authorize(creds)

    merchent_list_sheet = client.open('MTD OR').worksheet('merchent_list')
    merchent_list_data = merchent_list_sheet.get_all_values()
    merchent_list_headers = merchent_list_data.pop(0)
    merchent_list = pd.DataFrame(merchent_list_data, columns=merchent_list_headers)

    Tm_leader_target_sheet = client.open('MTD OR').worksheet('Tm_leader_target')
    Tm_leader_target_data = Tm_leader_target_sheet.get_all_values()
    Tm_leader_target_headers = Tm_leader_target_data.pop(0)
    tm_targ_df = pd.DataFrame(Tm_leader_target_data,columns=Tm_leader_target_headers)

    acc_mang_targ_sheet = client.open('MTD OR').worksheet('Tm_leader_target')
    acc_mang_targ_data = acc_mang_targ_sheet.get_all_values()
    acc_mang_targ_headers = acc_mang_targ_data.pop(0)
    acc_mang_targ_df = pd.DataFrame(acc_mang_targ_data,columns=acc_mang_targ_headers)

    return [tm_targ_df, merchent_list,acc_mang_targ_df]
st.set_page_config(page_title = 'Data admin App',layout='wide')

with st.sidebar :
    selected_menu = option_menu(
        menu_title = None,
        options= ['Achevment','Requested','Claim','Complaint']
    )
    




if selected_menu == 'Achevment' :
    Audf = False
    cudf = False

    with st.container():
        card,app = st.columns((1,1))
        
        with card:
            f_path = st.file_uploader('Upload card CSV',accept_multiple_files=False)
        with app :
            a_path = st.file_uploader('Upload App CSV',accept_multiple_files=False)
        dashBoards = st.checkbox('upload Data frame To DashBoards')
        set_report = st.button('processing')  
        if set_report :
            if f_path == None or f_path.name[-4:] != '.csv':
                with card:
                    st.error('must upload card as csv file')
            else:
                Audf = True
            if a_path == None or a_path.name[-4:] != '.csv':
                with app:
                    st.error('must upload csv file')
            else:
                cudf = True
            

        st.write('---')

    with st.container():
        if Audf == True and cudf == True :
            # -- variables to process the data --
            now = datetime.now()
            last_month = now.month - 1 if now.month > 1 else 12
            last_year = now.year if now.month > 1 else now.year - 1
            s_date = datetime(last_year, last_month, 1) + timedelta(days=32)
            s_date = s_date.replace(day=1) - timedelta(days=1)
            s_date = s_date.strftime("%Y-%m-%d")

            e_date = datetime(now.year, now.month, 1) + timedelta(days=32)
            e_date = e_date.replace(day=1) - timedelta(days=1)
            e_date = e_date.strftime("%Y-%m-%d")

            Today = date.date.today().strftime("%Y-%m-%d")
            Today = pd.to_datetime(Today)
            yesterday =Today - timedelta(days = 1)
            yesterday = yesterday.strftime("%Y-%m-%d")
            No_of_day = int(date.date.today().strftime("%d"))
            plan_day = int(cln.monthrange(date.datetime.now().year, date.datetime.now().month)[1])
            year = int(date.date.today().strftime("%Y"))
            tm_targ_df = get_processing_dfs()[0]
            merchent_list = get_processing_dfs()[1]
            acc_mang_targ_df = get_processing_dfs()[2]
            main_df = pd.read_csv(f_path, encoding='utf-8', encoding_errors='ignore', index_col=None)
            App_df = pd.read_csv(a_path, encoding='utf-8', encoding_errors='ignore', index_col=None)
            
            # Set Quarter data frame
            Quarter_data = {'start_date': [F'{year-1}-12-31', f'{year}-03-31', f'{year}-06-30', f'{year}-09-30'],
                    'end_date': [f'{year}-03-30', f'{year}-06-29', f'{year}-09-29', f'{year}-12-30'],
                    'Quarter': [f'Q1-{year}',f'Q2-{year}',f'Q3-{year}',f'Q4-{year}']}
            Quarter_df = pd.DataFrame(Quarter_data)
            Quarter_df['start_date'] = pd.to_datetime(Quarter_df['start_date'])
            Quarter_df['end_date'] = pd.to_datetime(Quarter_df['end_date'])
                        
            # Add team leader and **account manager** information
            main_df.rename(columns = {'Store_Merchant1' : 'Acount'}, inplace = True)
            App_df.rename(columns = {'DS_Name' : 'Acount'}, inplace = True)
            proccessing_data = pd.merge(main_df,merchent_list)
            App_df = pd.merge(App_df,merchent_list)

            # Filter active data in date range
            App_df = App_df[App_df['InstallmentStatus_Name'] == ('Active')]
            proccessing_data = proccessing_data[proccessing_data['TransactionType1'] == ('Active')]
            App_df['Installment_ModifiedDate'] = pd.to_datetime(App_df['Installment_ModifiedDate'])
            proccessing_data['ID_CreatedDate'] = pd.to_datetime(proccessing_data['ID_CreatedDate'])
            App_df = App_df[(App_df['Installment_ModifiedDate']>=(s_date)) & (App_df['Installment_ModifiedDate']<(e_date))]
            proccessing_data = proccessing_data[(proccessing_data['ID_CreatedDate']>=(s_date)) & (proccessing_data['ID_CreatedDate']<(e_date))]

            # Calculates the total achievement for each account
            achievement_per_account = proccessing_data.groupby(['Acount'])[['Total_Sales_without_AdminFees']].sum()
            achievement_per_account_app = App_df.groupby(['Acount'])[['Installmet_TotalAmount']].sum()
            achievement_per_account_df = pd.merge(achievement_per_account,achievement_per_account_app ,right_index = True , left_index = True, how = 'outer').fillna(0)
            achievement_per_account_df.rename(columns = {'Total_Sales_without_AdminFees' : 'Card'}, inplace = True)
            if 'Installmet_TotalAmount' in achievement_per_account_df.columns:
                achievement_per_account_df.rename(columns = {'Installmet_TotalAmount' : 'App'}, inplace = True)
            else:
                achievement_per_account_df['App'] = 0
            achievement_per_account_df["Total Achievement"] = achievement_per_account_df['App'] + achievement_per_account_df['Card']
            achievement_per_account_df.loc["Total Achievement"] = achievement_per_account_df.sum()
            achievement_per_account_df['Share %'] = achievement_per_account_df['Total Achievement'] / achievement_per_account_df['Total Achievement'].loc["Total Achievement"]
            achievement_per_account_df['Share %'] = np.round(achievement_per_account_df['Share %'], 4)
            achievement_per_account_df.sort_values("Total Achievement", axis = 0, ascending = False,
                            inplace = True, na_position ='last')
            achievement_per_account_df = pd.merge(achievement_per_account_df,merchent_list,left_on='Acount',right_on = 'Acount')
           
            # Calculates the total achievement for each account manager
            achievement_per_acc_manger =  proccessing_data.pivot_table(index =['Acount_manger'] ,values = ['Total_Sales_without_AdminFees'], aggfunc='sum')
            achievement_per_acc_manger_app  = App_df.pivot_table(index =['Acount_manger'] ,values = ['Installmet_TotalAmount'], aggfunc='sum')
            achievement_per_acc_manger_df = pd.merge(achievement_per_acc_manger,achievement_per_acc_manger_app ,right_index = True , left_index = True, how = 'outer').fillna(0)
            achievement_per_acc_manger_df.rename(columns = {'Total_Sales_without_AdminFees' : 'Card'}, inplace = True)

            if 'Installmet_TotalAmount' in achievement_per_acc_manger_df.columns:
                achievement_per_acc_manger_df.rename(columns = {'Installmet_TotalAmount' : 'App'}, inplace = True)
            else:
                achievement_per_acc_manger_df['App'] = 0
                
            achievement_per_acc_manger_df["Total Achievement"] = achievement_per_acc_manger_df['App'] + achievement_per_acc_manger_df['Card']
            achievement_per_acc_manger_df['Projection'] = achievement_per_acc_manger_df['Total Achievement']/No_of_day
            achievement_per_acc_manger_df['Projection'] = achievement_per_acc_manger_df['Projection'] * plan_day
            
            # Calculates the total achievement for each team leader
            achievement_per_team_leader =  proccessing_data.pivot_table(index =['Team_leader'] ,values = ['Total_Sales_without_AdminFees'], aggfunc='sum')
            achievement_per_team_leader_app  = App_df.pivot_table(index =['Team_leader'] ,values = ['Installmet_TotalAmount'], aggfunc='sum')
            achievement_per_team_leader_df = pd.merge(achievement_per_team_leader,achievement_per_team_leader_app ,right_index = True , left_index = True, how = 'outer').fillna(0)
            achievement_per_team_leader_df.rename(columns = {'Total_Sales_without_AdminFees' : 'Card'}, inplace = True)

            if 'Installmet_TotalAmount' in achievement_per_team_leader_df.columns:
                achievement_per_team_leader_df.rename(columns = {'Installmet_TotalAmount' : 'App'}, inplace = True)
            else:
                achievement_per_team_leader_df['App'] = 0
                
            achievement_per_team_leader_df["Total Achievement"] = achievement_per_team_leader_df['App'] + achievement_per_team_leader_df['Card']
            
            # Calculates the projection of total achievement

            achievement_per_team_leader_df['Projection'] = achievement_per_team_leader_df['Total Achievement']/No_of_day
            achievement_per_team_leader_df['Projection'] = achievement_per_team_leader_df['Projection'] * plan_day
            achievement_per_team_leader_df = pd.merge(tm_targ_df,achievement_per_team_leader_df,left_on='Team_leader',right_on='Team_leader')
            achievement_per_team_leader_df = achievement_per_team_leader_df.reindex(columns=['Team_leader','Card','App','Total Achievement','Target_Plan','Projection'])

            achievement_per_team_leader_df['Target_Plan'] = achievement_per_team_leader_df['Target_Plan'].astype(float)
            # calculate the projection ratio by dividing the projection by the target plan
            achievement_per_team_leader_df['Projection'] = achievement_per_team_leader_df['Projection'].astype(int)
            achievement_per_team_leader_df['Projection %'] = achievement_per_team_leader_df['Projection']/achievement_per_team_leader_df['Target_Plan']
            
            achievement_per_team_leader_df['Projection %'] = np.round(achievement_per_team_leader_df['Projection %'],4)

            # Add the projection ratio by dividing the projection by the target plan
            achievement_per_team_leader_df['Daily achievement rate'] = np.round((achievement_per_team_leader_df['Total Achievement'] / No_of_day),0).astype(int)
            
            # Calculates the required daily achievement by dividing the remaining target achievement by the remaining days
            if (plan_day - No_of_day) > 0 :
                achievement_per_team_leader_df['Required daily achievement'] = np.round((achievement_per_team_leader_df['Target_Plan'] - achievement_per_team_leader_df['Total Achievement'])/(plan_day - No_of_day),0).astype(int)
                for index, row in achievement_per_team_leader_df.iterrows():
                    if row['Required daily achievement'] < 0:
                        achievement_per_team_leader_df.at[index, 'Required daily achievement'] = 0
            else :
                achievement_per_team_leader_df['Required daily achievement'] = 0

            # *Integration of card and installment operations as one data frame*
            proccessing_data['Type'] = 'Card Transaction'
            App_df['Type'] = 'Normal App'

            #Select the specified columns
            proccessing_data = proccessing_data[['Acount', 'TransactionNumber1', 'ID_CreatedDate', 'Total_Sales_without_AdminFees', 'promoter', 'Acount_manger', 'Team_leader', 'Type']].copy()
            App_df = App_df[['Acount', 'Installment_UniqueID', 'Installment_ModifiedDate', 'Installmet_TotalAmount', 'promoter', 'Acount_manger', 'Team_leader', 'Type']]

            #Rename the columns with the appropriate names
            proccessing_data = proccessing_data.rename(columns={'TransactionNumber1': 'Transaction Number', 'ID_CreatedDate': 'Date', 'Total_Sales_without_AdminFees': 'Amount'})
            App_df = App_df.rename(columns={'Installment_UniqueID': 'Transaction Number', 'Installment_ModifiedDate': 'Date', 'Installmet_TotalAmount': 'Amount'})

            #Concatenate into one Dataframe defined as name **"All Achievement"**
            All_Achevment = pd.concat([proccessing_data, App_df])
            All_Achevment['Date'] = pd.to_datetime(All_Achevment['Date'])
            All_Achevment['month'] = pd.to_datetime(All_Achevment['Date']).dt.month_name()
            All_Achevment['Quarter'] = All_Achevment['Date'].apply(lambda x: Quarter_df[(Quarter_df['start_date'] <= x) & (Quarter_df['end_date'] >= x)]['Quarter'].values[0])

            animation_df = All_Achevment.copy()
            animation_df['Date'] = pd.to_datetime(animation_df['Date'])
            animation_df['Date'] = animation_df['Date'].apply(lambda x: x.strftime('%Y-%m-%d'))
            animation_achevment = animation_df.groupby(['Date','Acount'])[['Amount']].sum().astype(int)
            animation_achevment.reset_index(inplace=True)
            # Get unique dates and accounts from the original DataFrame
            daaate = animation_achevment['Date'].unique()
            acccount = animation_achevment['Acount'].unique()

            # Create a new DataFrame with columns for 'Account' and 'Date'
            animation_achevment_df = pd.DataFrame(columns=['Acount', 'Date'])

            # Loop over accounts and append a new row for each account with all unique dates
            for acc in acccount:
                for date in daaate:
                    row = {'Acount': acc, 'Date': date}
                    animation_achevment_df = pd.concat([animation_achevment_df, pd.DataFrame([row])], ignore_index=True)

            for d in range(len(daaate)):
                for a in range(len(acccount)):
                    anmit_amount = animation_df.loc[(animation_df['Acount'] == acccount[a]) & (animation_df['Date'] <= daaate[d]),'Amount'].sum()
                    animation_achevment_df.loc[(animation_achevment_df['Date']== daaate[d]) & (animation_achevment_df['Acount']== acccount[a]), 'Amount'] = anmit_amount


            if dashBoards :
                # Define the required scope of permissions.
                scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
                # permissions and authentication data  
                creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
                client = gspread.authorize(creds)

                # Open **"All_Achevmentsheet","per account",** and **"per team leader"** worksheet
                All_Achevmentsheet = client.open('MTD OR').worksheet('All_Achevment')
                per_accountsheet = client.open('MTD OR').worksheet('per account')
                per_per_teamleader = client.open('MTD OR').worksheet('per team leader')
                animation_achevment_sh = client.open('MTD OR').worksheet('animation_achevment_sh')
                # Clear existing data in worksheet
                All_Achevmentsheet.clear()
                per_accountsheet.clear()
                per_per_teamleader.clear()
                animation_achevment_sh.clear()
                print('Worksheet has cleared')

                # Update table data in worksheet
                All_Achevment['Date'] = All_Achevment['Date'].apply(lambda x: x.strftime('%Y-%m-%d %H:%M:%S'))
                per_accountsheet.update([achievement_per_account_df.columns.values.tolist()] + achievement_per_account_df.values.tolist())
                per_per_teamleader.update([achievement_per_team_leader_df.columns.values.tolist()] + achievement_per_team_leader_df.values.tolist())
                All_Achevmentsheet.update([All_Achevment.columns.values.tolist()] + All_Achevment.values.tolist())

                st.success("Data upladed To DashBoard")
                st.write('[manger DashBoard](https://lookerstudio.google.com/reporting/c7577406-f6ca-4bbb-885b-28ec59ded08e/page/IECLD)')
                st.write('[Leader DashBoard](https://lookerstudio.google.com/reporting/2c97d361-2d2d-4b4f-81bc-8d0691736381)')
            else :
                #Display on app 

                with st.container():
                    
                    st.subheader('All Active Transaction')
                    
                    
                    st.write('##')
                    st.dataframe(All_Achevment)
                    
                    
                    
                    

                    
                    
                    
                        
                    

                    writer = pd.ExcelWriter('All_Achevment.xlsx')

                
                    All_Achevment.to_excel(writer, sheet_name='All_Achevment', index=False)
                    
                    writer.close()

                    # إنشاء رابط التنزيل
                    with open('All_Achevment.xlsx', 'rb') as f:
                        data = f.read()
                    b64 = base64.b64encode(data).decode('utf-8')
                    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="All_Achevment.xlsx"> Export as Excel file </a>'
                    st.markdown(href, unsafe_allow_html=True)
                    st.write('---')
                with st.container():
                    st.subheader('achievement per account Table')
                    st.write('##')
                    #All_Achevment_coulmuns = st.columns((1,2))
                    #with All_Achevment_coulmuns[0]:
                        #st.dataframe(achievement_per_account_df)
                    #with All_Achevment_coulmuns[1]:
                    raceplot = barplot(animation_achevment_df,  item_column="Acount", value_column="Amount", time_column="Date")
                    fig = raceplot.plot(item_label='All_Achevment', value_label="EGP", frame_duration=800, date_format='%Y-%m-%d', orientation='horizontal')
                    fig.update_layout(
                    title='opration',
                    autosize=False,
                    width=1000,
                    height=500  ,
                    paper_bgcolor="lightgray"
                    )
                    st.plotly_chart(fig)
                    
        
                    
                st.write('---')

                
                with st.container():
                    st.subheader('achievement per team leader df')
                    st.write('##')
                    st.dataframe(achievement_per_team_leader_df)
                st.write('---')
                All_Achevment['Date']=pd.to_datetime(All_Achevment['Date'])
                yesterday_Achevment = All_Achevment[(All_Achevment['Date']>(yesterday)) & (All_Achevment['Date']<(Today))].copy()
                yesterday_team = yesterday_Achevment.groupby(['Team_leader'])[['Amount']].sum()
                yest_Achevment = yesterday_Achevment['Amount'].sum()
                with st.container():
                    st.subheader(f'yesterday Achevment : {yest_Achevment}' )
                    st.write('##')
                    st.dataframe(yesterday_team)
                st.write('---')

            

if selected_menu == 'Requested' :

    with st.container():
        Arudf = False
        crudf = False
        card,app = st.columns((1,1))
        
        with card:
            Card_App_path = st.file_uploader('Upload card request CSV',accept_multiple_files=False)
        with app :
            Normal_App_path = st.file_uploader('Upload App request CSV',accept_multiple_files=False)
        set_report = st.button('processing')  
        if set_report :
            if Card_App_path == None or Normal_App_path.name[-4:] != '.csv':
                with card:
                    st.error('must upload card as csv file')
            else:
                Arudf = True
            if Normal_App_path == None or Normal_App_path.name[-4:] != '.csv':
                with app:
                    st.error('must upload csv file')
            else:
                crudf = True
            

        st.write('---')
    with st.container():
        if Arudf == True and crudf == True :
            # -date variables-
            Today = date.date.today().strftime("%Y-%m-%d")
            Today = pd.to_datetime(Today)
            yesterday =Today - timedelta(days = 1)
            yesterday = yesterday.strftime("%Y-%m-%d")

            # Load data
            merchent_list = get_processing_dfs()[1]
            Card_App_df = pd.read_csv(Card_App_path, encoding='utf-8', encoding_errors='ignore')
            Normal_App_df = pd.read_csv(Normal_App_path)
            #Add **team leader** and **account manager** information
            Card_App_df.rename(columns = {'Store_Name' : 'Acount'}, inplace = True)
            Normal_App_df.rename(columns = {'DS_Name' : 'Acount'}, inplace = True)
            Card_App_df = pd.merge(Card_App_df,merchent_list  )
            Normal_App_df = pd.merge(Normal_App_df,merchent_list )
            #Select the specified columns
            Normal_Request = Normal_App_df[['D_Name1','Acount','Installment_UniqueID','Installment_CreatedDate','Installment_ModifiedDate','Installmet_TotalAmount','InstallmentStatus_Name','Team_leader','Comment','RejectionReason']].copy()
            Card_Request = Card_App_df[['D_Name1','Acount','Installment_UniqueID','Installment_CreatedDate','Installment_ModifiedDate','Installment_Amount','InstallmentStatus_Name','Team_leader','Comment','CancellationReasons']].copy()
            #Rename the columns with the appropriate names
            Normal_Request.rename(columns = {'Installment_UniqueID' : 'Request number','Installment_CreatedDate' : 'Created Date','Installment_ModifiedDate' : 'Modified Date','Installmet_TotalAmount' : 'Installmet Amount'}, inplace = True)
            Card_Request.rename(columns = {'Installment_UniqueID' : 'Request number','Installment_CreatedDate' : 'Created Date','Installment_ModifiedDate' : 'Modified Date','Installment_Amount' : 'Installmet Amount'}, inplace = True)
            #Set the the **"Type"**  column
            Card_App_df['App type'] = 'Card Request'
            Normal_App_df['App type'] = 'Normal Request'
            #Concatenate into one Dataframe defined as name **"All_Request"**
            All_Request = pd.concat([Normal_Request,Card_Request])
            #preparing Date as datetime & Set the  **"month_creation"**  column
            All_Request['Created Date'] = pd.to_datetime(All_Request['Created Date'])
            All_Request['month_creation'] = All_Request['Created Date'].dt.month_name()
            
            st.header('proccessing data')
            st.dataframe(All_Request)
            

            All_Request_excel_file = 'All_Request.xlsx'
            All_Request.to_excel(All_Request_excel_file, index=False)

            # تحميل ملف Excel باستخدام Streamlit
            with open(All_Request_excel_file, 'rb') as f:
                excel_bytes = f.read()
                b64 = base64.b64encode(excel_bytes).decode()
                href = f'<a href="data:application/octet-stream;base64,{b64}" download="{All_Request_excel_file}">Download Excel File</a>'
                st.markdown(href, unsafe_allow_html=True)
                st.write('---')


if selected_menu == 'Claim' :


    with st.container():
        Cludf = False
        
        Claim,space = st.columns((1,1))
        
        with Claim:
            claim_path = st.file_uploader('Upload Claim request CSV',accept_multiple_files=False)
        
        set_report = st.button('processing')  
        if set_report :
            if claim_path == None or claim_path.name[-4:] != '.csv':
                with Claim:
                    st.error('must upload claims as csv file')
            else:
                Cludf = True
    st.write('---')
    with st.container():
        if Cludf == True :
            claim_df = pd.read_csv(claim_path)
            merchent_list_df = get_processing_dfs()[1]
            
            ## Data cleaning and manipulation
            if isinstance(claim_df, pd.DataFrame) & isinstance(merchent_list_df, pd.DataFrame):
                merchent_list_df.rename(columns = {'Acount' : 'Store_Name'}, inplace = True)
                pros_claim = pd.merge(claim_df,merchent_list_df)
                pros_claim['Claim_CreatedDate'] = pd.to_datetime(pros_claim['Claim_CreatedDate'])
                pros_claim['Claim_ModifiedDate'] = pd.to_datetime(pros_claim['Claim_ModifiedDate'])
                pros_claim['today'] = pd.to_datetime(date.date.today().strftime("%Y-%m-%d"))
                pros_claim['Modifai_SLA'] = (pros_claim['today'] - pros_claim['Claim_ModifiedDate']).dt.days + 1
                pros_claim['Create_SLA'] = (pros_claim['today'] - pros_claim['Claim_CreatedDate']).dt.days
                pros_claim['close_SLA'] = (pros_claim['Claim_ModifiedDate'] - pros_claim['Claim_CreatedDate']).dt.days
                pros_claim['Net_Claim_Amount'].fillna(0, inplace=True)
                pros_claim = pros_claim[['Claim_id','Net_Claim_Amount','ClaimStatus_Name','Store_Name','Claim_CreatedBy','Claim_CreatedDate','Claim_ModifiedDate','Team_leader','close_SLA','Modifai_SLA','Create_SLA']]
                pros_claim['Net_Claim_Amount'] = pd.to_numeric(pros_claim['Net_Claim_Amount'], errors='coerce').dropna().astype(int)

            ## Filter data by claim status
            pending_AR = pros_claim[pros_claim['ClaimStatus_Name']=='Pending  Ar']
            AR_Rejected = pros_claim[pros_claim['ClaimStatus_Name']=='AR Rejected']
            Pending_Payable = pros_claim[pros_claim['ClaimStatus_Name']=='Pending  Payable']
            Pending_Treasurment = pros_claim[pros_claim['ClaimStatus_Name']=='Pending Treasurment']
            New = pros_claim[pros_claim['ClaimStatus_Name']=='New']
            Close = pros_claim[pros_claim['ClaimStatus_Name']=='Close']
            Today = pd.to_datetime(date.date.today().strftime("%Y-%m-%d"))
            yesterday = (Today - date.timedelta(days=1)).strftime("%Y-%m-%d")
            Close_Today = Close[(Close['Claim_ModifiedDate'] >= (Today))]
            Close_yesterDay = Close[(Close['Claim_ModifiedDate'] >= yesterday) & (Close['Claim_ModifiedDate'] <= Today)]

            ## Create summary table
            Claim_Review = pd.DataFrame({'status' :['Close_Today', 'Close_yesterDay', 'pending_AR', 'AR_Rejected', 'Pending_Payable', 'Pending_Treasurment'],
                'Count' : [Close_Today['Claim_ModifiedDate'].count(), Close_yesterDay['Claim_ModifiedDate'].count(),
            pending_AR['Claim_ModifiedDate'].count(), AR_Rejected['Claim_ModifiedDate'].count(),
            Pending_Payable['Claim_ModifiedDate'].count(), Pending_Treasurment['Claim_ModifiedDate'].count()],
                'Claim_Amount' : [Close_Today['Net_Claim_Amount'].sum(), Close_yesterDay['Net_Claim_Amount'].sum(),
            pending_AR['Net_Claim_Amount'].sum(), AR_Rejected['Net_Claim_Amount'].sum(),
            Pending_Payable['Net_Claim_Amount'].sum(), Pending_Treasurment['Net_Claim_Amount'].sum()],
                'Arage_SLA' : [Close_Today['Modifai_SLA'].mean(), Close_yesterDay['Modifai_SLA'].mean(),
            pending_AR['Modifai_SLA'].mean(), AR_Rejected['Modifai_SLA'].mean(),
            Pending_Payable['Modifai_SLA'].mean(), Pending_Treasurment['Modifai_SLA'].mean()]})
            Claim_Review['Arage_SLA'] = Claim_Review['Arage_SLA'].fillna(0).apply(np.ceil).astype(int)


            st.dataframe(Claim_Review)

            New_Claims,pending_AR_Claims,AR_Rejected_Claims,Pending_Payable_Claims,Pending_Treasurment_Claims,Close_Today_Claims,Close_yesterDay_Claims = st.tabs([
                "New","pending AR", "AR Rejected",
                "Pending Payable",'Pending Treasurment',
                'Close Today', 'Close yesterDay'
                ])
            
            New_Claims.dataframe(New)
            pending_AR_Claims.dataframe(pending_AR)
            AR_Rejected_Claims.dataframe(AR_Rejected)
            Pending_Payable_Claims.dataframe(Pending_Payable)
            Pending_Treasurment_Claims.dataframe(Pending_Treasurment)
            Close_Today_Claims.dataframe(Close_Today)
            Close_yesterDay_Claims.dataframe(Close_yesterDay)

            
            
           
            writer = pd.ExcelWriter('claim_review.xlsx')

            
            Claim_Review.to_excel(writer, sheet_name='Claim_Review', index=False)
            New.to_excel(writer, sheet_name='New', index=False)
            pending_AR.to_excel(writer, sheet_name='Pending_AR', index=False)
            AR_Rejected.to_excel(writer, sheet_name='AR_Rejected', index=False)
            Pending_Payable.to_excel(writer, sheet_name='Pending_Payable', index=False)
            Pending_Treasurment.to_excel(writer, sheet_name='Pending_Treasurment', index=False)
            Close_Today.to_excel(writer, sheet_name='Close_Today', index=False)
            Close_yesterDay.to_excel(writer, sheet_name='Close_yesterDay', index=False)

            writer.close()

            # إنشاء رابط التنزيل
            with open('claim_review.xlsx', 'rb') as f:
                data = f.read()
            b64 = base64.b64encode(data).decode('utf-8')
            href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="claim_review.xlsx"> Export as Excel file </a>'
            st.markdown(href, unsafe_allow_html=True)


if selected_menu == 'Complaint' :


    with st.container():
        Coudf = False
        
        complaint,space = st.columns((1,1))
        
        with complaint:
            complaint_path = st.file_uploader('Upload card complaint xlsx',accept_multiple_files=False)
        
        set_report = st.button('processing')  
        if set_report :
            if complaint_path == None or complaint_path.name[-5:] != '.xlsx':
                with complaint:
                    st.error('must upload complaint as xlsx file')
            else:
                Coudf = True
    st.write('---')
    if Coudf == True:
        # processing  Data 
        Today = date.date.today().strftime("%Y-%m-%d")
        # Load data
        complaint_data = pd.read_excel(complaint_path).rename(columns={'الفرع': 'Acount'})
        merchent_list = get_processing_dfs()[1]
        # Merge data and export to Excel
        complaint_data = pd.merge(complaint_data, merchent_list)
        # Brief of complaint
        complaint_pivot = complaint_data.pivot_table(columns=['Team_leader'], values=['رقم_الشكوى_الإستعلام'], aggfunc='count').rename(mapper={'رقم_الشكوى_الإستعلام': 'Count of complaints'})

        with st.container():
            st.subheader('Brief of complaint')
            st.table(complaint_pivot)
            st.write('---')
            team_list = ['All Teams']
            team_list.extend(complaint_data['Team_leader'].unique())
            
            #locals()['aaaa'] = None
            
            complaint_taps = st.tabs(team_list)
            complaint_taps[0].table(complaint_data)
            complaint_TI = 1
            
            for i in range(len(team_list)-1):
                complaint_taps[complaint_TI].table(complaint_data.loc[complaint_data['Team_leader'] == team_list[complaint_TI]])
                complaint_TI+=1


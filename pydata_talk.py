import xlwings as xw
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# xlwings UDFs -----------------------------------------

@xw.func
@xw.arg('df', pd.DataFrame, index=False, header=True)
def issue_status(df):
    return df['state'].value_counts()

@xw.func
@xw.arg('df', pd.DataFrame, index=False, header=True)
def top_issues_labels(df):
    return (df['labels'].str.split(', ')
                               .explode()
                               .value_counts()
                               .head())

@xw.func
@xw.arg('df', pd.DataFrame, index=False, header=True)
def issues_created_resolved(df):
    closed_issues_df = df[df['state'] == 'closed']
    
    df['created_at'] = pd.to_datetime(df['created_at'])
    df['updated_at'] = pd.to_datetime(df['updated_at'])
    
    issues_resolved_per_day = (closed_issues_df.groupby(df['updated_at'].dt.date)
                               .size()
                               .reset_index(name='resolved_count'))

    issues_resolved_per_day.rename(columns={'updated_at': 'date'}, inplace=True)

    # Combine created and resolved issues per day
    issues_created_per_day = (df.groupby(df['created_at'].dt.date)
                              .size()
                              .reset_index(name='created_count'))

    issues_created_per_day.rename(columns={'created_at': 'date'}, inplace=True)

    combined_issues_daily = (pd.merge(issues_created_per_day, 
                                      issues_resolved_per_day, 
                                      on='date', how='outer')
                             .fillna(0))
    combined_issues_daily['date'] = (combined_issues_daily['date']
                                     .map(lambda d: d.strftime("%d-%m-%Y")))
    return combined_issues_daily

@xw.func
@xw.arg('df', pd.DataFrame, index=False, header=True)
def resol_eff(df):
    return df.describe().astype(int)


@xw.func
@xw.arg('df', pd.DataFrame, index=False, header=True)
def resol_eff_chart(df, caller):
    fig = plt.figure(figsize=(10, 5))
    sns.lineplot(x='date', y='created_count', 
                 data=df, 
                 label='Issues Created', color = '#D9D9D9')

    sns.lineplot(x='date', y='resolved_count', 
                 data=df, 
                 label='Issues Resolved', color = '#76933C')
    plt.title('Issues Created vs. Issues Resolved', 
              loc= 'left', fontsize = 16, fontweight='bold', pad=25)

    plt.xlabel(None)
    plt.ylabel('Number of Issues')
    plt.xticks(rotation=45)
    plt.legend(frameon=False)
    plt.grid(False)
    plt.ylim([0, 27])
    plt.tight_layout()
    sns.despine()
    ws= xw.Book().sheets[0]
    caller.sheet.pictures.add(fig, name='resol_eff_chart', update=True,
                              left=ws.range('W2').left, 
                              top=ws.range('W2').top)
    
    return 'Plotted Resolution Efficiency Chart'

# CODE FOR Run Main-----------------------------------------

def load_data(sheet):
    return (sheet.range("A1")
          .options(pd.DataFrame, header=1, index=False, expand="table")
          .value)
    
def issue_count(df):
    return df['state'].value_counts()

def top_issues_by_labels(df):
    return (df['labels'].str.split(', ')
                               .explode()
                               .value_counts()
                               .head())

def issues_created_and_resolved(df):
    closed_issues_df = df[df['state'] == 'closed']
    
    df['created_at'] = pd.to_datetime(df['created_at'])
    df['updated_at'] = pd.to_datetime(df['updated_at'])
    
    issues_resolved_per_day = (closed_issues_df.groupby(df['updated_at'].dt.date)
                               .size()
                               .reset_index(name='resolved_count'))

    issues_resolved_per_day.rename(columns={'updated_at': 'date'}, inplace=True)

    # Combine created and resolved issues per day
    issues_created_per_day = (df.groupby(df['created_at'].dt.date)
                              .size()
                              .reset_index(name='created_count'))

    issues_created_per_day.rename(columns={'created_at': 'date'}, inplace=True)

    combined_issues_daily = (pd.merge(issues_created_per_day, 
                                      issues_resolved_per_day, 
                                      on='date', how='outer')
                             .fillna(0))
    
    return combined_issues_daily.describe().astype(int)

def write_results(sheet, df):
    sheet.range('A1').value = "Issue Status" 
    sheet.range('A3').value = issue_count(df)
    
    sheet.range('D1').value = 'Top Issue Labels'
    sheet.range('D3').value = top_issues_by_labels(df)
    
    
    sheet.range('G1').value = 'Stats_ Issues created and resolved per day'
    sheet.range('G3').value = issues_created_and_resolved(df)
    
    
def formatting(sheet):
    sheet.range('A1:B1').merge()
    sheet.range('D1:E1').merge()
    sheet.range('G1:J1').merge()
 
    for cell in ['A1', 'D1', 'G1']:
        sheet.range(cell).font.bold = True
        sheet.range(cell).color = '#CAEDFB'
        
    sheet.autofit(axis='columns')
     
  

def main():
    wb = xw.Book.caller()       # Connect to Excel
    ws1 = wb.sheets[0]
    df = load_data(ws1)       # Load data
    
    try:
        ws2 = wb.sheets['Analysis']
        ws2.clear() # clear old summary
    except:
        ws2 = wb.sheets.add('Analysis', after='Sheet1')
    write_results(ws2, df) # Write results 
    formatting(ws2)           # Format Sheet    
            
    
# Main Guard __________________________________________________________________   
    
if __name__ == '__main__':
    xw.Book("demo.xlsm").set_mock_caller()
    main()
    


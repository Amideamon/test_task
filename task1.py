import pandas as pd

df_test = pd.read_excel('test_input.xlsx')
filename = df_test.iloc[1][1]

df_test = pd.read_excel('test_input.xlsx', header=4)

with open('task1.xml', 'w+') as test_file:
    test_file.write('<?xml version="1.0" encoding="UTF-8"?>\n')
    test_file.write('<CERTDATA>\n')
    test_file.write('\t<FILENAME>' + filename +'</FILENAME>\n')
    
    test_file.write('\t<ENVELOPE>\n')
    
    for i in range(len(df_test)):
        test_file.write('\t\t<ECERT>\n')
        
        test_file.write(f'\t\t\t<CERTNO>{df_test.iloc[i]["Ref no"]}</CERTNO>\n')
        
        test_file.write(f'\t\t\t<CERTDATE>{str(df_test.iloc[i]["Issuance Date"].date())}</CERTDATE>\n')
        
        test_file.write(f'\t\t\t<STATUS>{df_test.iloc[i]["Status"]}</STATUS>\n')
        
        test_file.write(f'\t\t\t<IEC>0{df_test.iloc[i]["IE Code"]}</IEC>\n')
        
        test_file.write(f'\t\t\t<EXPNAME>"{df_test.iloc[i]["Client"]}"</EXPNAME>\n')
        
        test_file.write(f'\t\t\t<BILLID>{df_test.iloc[i]["Bill Ref no"]}</BILLID>\n')
        
        test_file.write(f'\t\t\t<SDATE>{str(df_test.iloc[i]["SB Date"].date())}</SDATE>\n')
        
        test_file.write(f'\t\t\t<SCC>{df_test.iloc[i]["SB Currency"]}</SCC>\n')
        
        sb_amount = df_test.iloc[i]['SB Amount']
        if str(sb_amount)[-2:] == '.0': 
            test_file.write(f'\t\t\t<SVALUE>{sb_amount:.0f}</SVALUE>\n')
        else: 
            test_file.write(f'\t\t\t<SVALUE>{sb_amount:.2f}</SVALUE>\n')
        
        test_file.write('\t\t</ECERT>\n')
        
    test_file.write('\t</ENVELOPE>\n')
    test_file.write('</CERTDATA>\n')
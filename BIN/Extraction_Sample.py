import SAPLogon as SAP
import os
if __name__ == "__main__":
    #SAP credencials
    user = "User_Fake"  #this is a fake user, you need to use your own user
    password = "Pass_fake"  #this is a fake password, you need to use your own password
    

    #define columns to extract  
    ReportsColumns = ["Description", "Budget","Proj.cost sched.000","Actual revenues","Actual costs",
           "Project rev.plan 000","Total cost commitment","Status"]

    
    #define path to save the files
    run_path=os.getcwd() #make sure that this is the main folder were is the BIN and CONFIG folders
    print(run_path)
    TempFilesPath=os.path.join(run_path, 'CONFIG')

    #define the project list, alll list are a list of strings
    pdlist=['EUS02.SW713']
    WBSlist=['EUS02.SW713.01.007',
             'EUS02.SW713.01.095',
             'EUS02.SW713.01.042',
             'EUS02.SW713.01.010',
             'EUS02.SW713.01.021',
             'EUS02.SW713.01.005',
             'EUS02.SW713.01.013',
             'EUS02.SW713.01.024',
             'EUS02.SW713.01.017',
             'EUS02.SW713.01.005']
    polist=['13992010','13992054','13992055']
    
    # Create a SAPLogon object
    sap = SAP.SapGui()
    #logon to SAP
    #As a recomendation, allow scripting in your SAP account 
    #and desactivate the popup to allow the script to run
    sap.sap_Login(pUserName=user,pPassword=password)


    #this layouts arready exist in my acount, you need use the layout that you have in your account or the default
    
    #sap.get_CNS41(PD=pdlist,ouput_Path=TempFilesPath,Columns=ReportsColumns)
    #sap.get_CJI3(PD=pdlist,ouput_Path=TempFilesPath,Layout="/CJI3_BASIC")
    #sap.get_CJI5(PD=pdlist,ouput_Path=TempFilesPath,Layout="/CMMITDETAIL")
    #sap.get_ZSNAP(customer_PO_number=polist,ouput_Path=TempFilesPath,Layout="/ZSNAP_R&V")
    #sap.get_Zzcustmon(WBS_element=WBSlist,ouput_Path=TempFilesPath,Layout="/ZZ_R&V")
    #sap.get_zcheckqtc(WBS_element=WBSlist,ouput_Path=TempFilesPath,Layout="/ZZCHECK_PNW")
    #sap.get_CN46N(PD=pdlist,ouput_Path=TempFilesPath)

    sap.sap_close()
    del sap
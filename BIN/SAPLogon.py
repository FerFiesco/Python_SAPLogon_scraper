import win32com.client
import sys
import subprocess
import time
from os.path import exists
from os import remove,path
from os import makedirs
from os.path import join
from datetime import date
from pathlib import Path

bin_path = path.join(path.dirname(path.realpath(sys.argv[0])))
config_path = path.join(bin_path[:-3], 'CONFIG')
if not path.isdir(config_path):   makedirs (config_path)
output_path = path.join(bin_path[:-3], 'OUTPUT')
if not path.isdir(output_path):   makedirs (output_path)
#Login into SAP
class SapGui():
        def __init__(self):
                self.extract_counter=0
                self.path =r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
                self.sap_subprocess = subprocess.Popen(self.path)
                time.sleep(5)

        #Select SAP GUI for manipulation
                self.SapGuiAuto = win32com.client.GetObject("SAPGUI")
                if not type(self.SapGuiAuto) == win32com.client.CDispatch:
                        print("SAP GUI not found")
                        return

        #Get scripting engine
                self.application = self.SapGuiAuto.GetScriptingEngine

        def sap_close(self):
                self.session.findById("wnd[0]").maximize()
                self.session.findById("wnd[0]").close()
                self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
                time.sleep(1)
                self.sap_subprocess.kill()
        def sap_Login(self,pUserName="",pPassword=""):
                try:
                        self.sap_Logon(pUserName,pPassword)
                        return True
                except:
                        print(sys.exc_info()[0])
                        #messagebox.showinfo('showinfo', 'login successfuly')
                        return False
                pass
        def validate_run_response(self) :
                # <----------------------------------Validate transaction running----------------------------------------------------------
                
                if self.session.Children.Count > 1 :
                        WINDOW_TITLE = self.session.findById("wnd[1]").text
                        LC_TEST_SUB_STRING_1 = "Information"
                        LC_WINDOW_CHECK = WINDOW_TITLE.find(LC_TEST_SUB_STRING_1)
                        if LC_WINDOW_CHECK >= 0:
                                SAPResponseMessage = self.session.findById("wnd[1]/usr/txtMESSTXT1").text
                                print(SAPResponseMessage)
                                self.session.findById("wnd[1]/tbar[0]/btn[0]").press
                                return False

                SAPResponseMessage = self.session.findById("wnd[0]/sbar").text
                if len(SAPResponseMessage) > 0 :
                        if SAPResponseMessage.find("Objects were filtered") >= 0:
                                print(SAPResponseMessage)
                                return True
                        elif SAPResponseMessage.find('Name or password is incorrect (repeat logon)') >= 0:
                                print(SAPResponseMessage)
                                self.session.findById("wnd[0]/tbar[0]/btn[15]").press()
                                return False
                        elif SAPResponseMessage.find('Selection canceled after 5,000 data records') >= 0:
                                print(SAPResponseMessage)
                                self.session.findById("wnd[0]/tbar[0]/btn[15]").press()
                                return False
                        elif SAPResponseMessage.find('Memory low. Leave the transaction before taking a break!') >= 0:
                                print(SAPResponseMessage)
                                self.session.findById("wnd[0]/tbar[0]/btn[15]").press()
                                self.session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
                                return False
                        elif SAPResponseMessage.find('No line items were selected') >= 0:
                                print(SAPResponseMessage)
                                self.session.findById("wnd[0]/tbar[0]/btn[15]").press()
                                return False
                        elif SAPResponseMessage.find('Memory low. Leave the transaction before taking a break!') >= 0:
                                print(SAPResponseMessage)
                                self.session.findById("wnd[0]/tbar[0]/btn[15]").press()
                                return False
                        elif SAPResponseMessage.find('No record found') >= 0:
                                print(SAPResponseMessage)
                                self.session.findById("wnd[0]/tbar[0]/btn[15]").press()
                                return False
                        elif SAPResponseMessage.find('Choose a valid function') >= 0:
                                print(SAPResponseMessage)
                                self.session.findById("wnd[0]/tbar[0]/btn[15]").press()
                                return False 
                        else:
                                print(SAPResponseMessage)
                                self.session.findById("wnd[0]/tbar[0]/btn[15]").press()
                                self.session.findById("wnd[0]/tbar[0]/btn[15]").press()
                                return False                  
                else:
                        print("Done")
                        return True

        def sap_Logon(self,pUserName="",pPassword=""):

                print("Login: " + pUserName)     
                #Stablish connection
                self.connection = self.application.OpenConnection("P12 ONE Password Logon",True)
                time.sleep(5)
                if not type(self.connection) == win32com.client.CDispatch:
                        application = None
                        SapGuiAuto = None
                        return False
                
        #Open connection
                self.session = self.connection.Children(0)
                if not type(self.session) == win32com.client.CDispatch:
                        print("SAP session not open")
                        connection = None
                        application = None
                        SapGuiAuto = None
                        return False
        
        #Logon PAssword       
                SAPErrorMessage  = "Name or password is incorrect (repeat logon)"
                
                self.session.findById("wnd[0]").SetFocus()
                self.session.findById("wnd[0]/usr/txtRSYST-LANGU").text = "EN"
                self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = pUserName
                self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = pPassword
                self.session.findById("wnd[0]/tbar[0]/btn[0]").SetFocus()
                
                self.session.findById("wnd[0]/tbar[0]/btn[0]").press()

                
                if self.session.Children.Count > 1 :
                        WINDOW_TITLE = self.session.findById("wnd[1]").text
                        LC_TEST_SUB_STRING_1 = "License Information for Multiple Logon"
                        LC_TEST_SUB_STRING_2 = "System Messages"
                        LC_WINDOW_CHECK_MULTIPLE_LOGON = WINDOW_TITLE.find(LC_TEST_SUB_STRING_1)
                        LC_WINDOW_CHECK_SYSTEM_MESSAGES = WINDOW_TITLE.find( LC_TEST_SUB_STRING_2)
                        if LC_WINDOW_CHECK_MULTIPLE_LOGON >= 0:
                                self.session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").Select()
                                self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                        
                        if LC_WINDOW_CHECK_SYSTEM_MESSAGES >= 0 :
                                self.session.findById("wnd[1]/tbar[0]/btn[12]").press()
              
                time.sleep(2)
                if self.session.Children.Count > 1 :
                        WINDOW_TITLE = self.session.findById("wnd[1]").text
                        LC_TEST_SUB_STRING_1 = "License Information for Multiple Logon"
                        LC_TEST_SUB_STRING_2 = "System Messages"
                        LC_WINDOW_CHECK_MULTIPLE_LOGON = WINDOW_TITLE.find(LC_TEST_SUB_STRING_1)
                        LC_WINDOW_CHECK_SYSTEM_MESSAGES = WINDOW_TITLE.find( LC_TEST_SUB_STRING_2)
                        if LC_WINDOW_CHECK_MULTIPLE_LOGON >= 0:
                                self.session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").Select()
                                self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                        
                        if LC_WINDOW_CHECK_SYSTEM_MESSAGES >= 0 :
                                self.session.findById("wnd[1]/tbar[0]/btn[12]").press()
                time.sleep(2)
                return self.validate_run_response()

        def set_transaction_parameter(self,id="",value_list=[""],ClipBoard_path=config_path):
                if value_list[0]!="":
                        text_file = open(join(ClipBoard_path,"ClipBoard.txt"), "w")
                        text_file.write('\n'.join(value_list))
                        text_file.close()

                        self.session.findById(id).press()
                        self.session.findById("wnd[1]/tbar[0]/btn[16]").press() #--Clean values
                        time.sleep(0.5)
                        self.session.findById("wnd[1]/tbar[0]/btn[23]").press()
                        time.sleep(1)
                        self.session.findById("wnd[2]/usr/ctxtDY_PATH").text = ClipBoard_path
                        self.session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = "ClipBoard.txt"
                        self.session.findById("wnd[2]/tbar[0]/btn[0]").press()

                        time.sleep(1)
                        self.session.findById("wnd[1]/tbar[0]/btn[8]").press() #-OK

        

        #----------------------------------------------------------------------------------CNS41---------------------------------------------------------------------------------------
        def get_CNS41(self, 
                PD=[""],
                Sales_document=[""],
                WBS_element=[""],
                Network_order = [""],
                Activity=[""],
                Materials_in_network=[""],
                level=[1,4],
                filename="",
                ouput_Path=config_path,
                Columns=[""]
                ):
                if PD[0]!="" or Sales_document[0]!="" or WBS_element[0]!="" or Network_order[0]!="" or Activity[0]!="" or Materials_in_network[0]!="":

                        if len(filename)<1:
                                default_filename="CNS41_" + PD[0] +"_"+ Sales_document[0] +"_"+ WBS_element[0] +"_"+ Network_order[0] +"_"+ Activity[0] +"_"+ Materials_in_network[0] +".txt"
                                default_filename = default_filename.replace("*","")
                                default_filename = default_filename.replace("__","_")
                                default_filename = default_filename.replace("__","_")
                                default_filename = default_filename.replace("__","_")
                                default_filename = default_filename.replace("_.txt",".txt")
                                filename=default_filename
                        else:
                                if filename.find(".") >= 0:
                                        filename=filename[0:filename.find(".")+1]+".txt"
                                else:
                                        filename=filename + ".txt"

                        if ouput_Path[-1]!="/":
                                ouput_Path = ouput_Path+"/"

                        file_exists = exists(ouput_Path + filename )
                        if file_exists :
                                remove(ouput_Path + filename)
                        #open transaction
                        self.session.findById("wnd[0]").maximize
                        self.session.findById("wnd[0]/tbar[0]/okcd").text = "CNS41"
                        self.session.findById("wnd[0]").sendVKey(0)

                        if self.session.Children.Count > 1 :
                                self.session.findById("wnd[1]/usr/ctxtTCNTT-PROFID").text = "Z00000000001"
                                self.session.findById("wnd[1]").sendVKey (0)
                        

                        #fill fields in transaction
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_CN_PROJN_%_APP_%-VALU_PUSH",value_list=PD,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_CN_VBELN_%_APP_%-VALU_PUSH",value_list=Sales_document,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_CN_PSPNR_%_APP_%-VALU_PUSH",value_list=WBS_element,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_CN_NETNR_%_APP_%-VALU_PUSH",value_list=Network_order,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_CN_ACTVT_%_APP_%-VALU_PUSH",value_list=Activity,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_CN_MATNR_%_APP_%-VALU_PUSH",value_list=Materials_in_network,ClipBoard_path=ouput_Path)
                        
                        #set level for extraction
                        self.session.findById("wnd[0]/usr/txtCN_STUFE-LOW").text = min(level)
                        self.session.findById("wnd[0]/usr/txtCN_STUFE-HIGH").text = max(level)
                        self.session.findById("wnd[0]").sendVKey (0)

                        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()  # Run  <--------------------
                        time.sleep(1)
                        # ----------------------------------Validate transaction running ------------------------------------------------------
                         
                        response=self.validate_run_response()
                        if not response: return response
                        time.sleep(3)


                        #--------------------------------------------------SET columns---------------------------------------------------
                        if Columns[0] != "" :   
                                self.session.findById("wnd[0]/tbar[1]/btn[5]").press()
                                self.session.findById("wnd[1]/usr/ssubRAHMEN:SAPLCNFA:0111/subAUSWAHL:SAPLCNFA:0140/btnALLE_NICHT_AUSWAEHLEN").press()
                                self.session.findById("wnd[1]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/tblSAPLCNFATC_ALLE_FELDER").verticalScrollbar.position = 0
                                selected_cols_coun=0
                                WhileCount = 0
                                WhileCondition = True
                                while WhileCondition: #Inserted While to ensure all columns are retrieved ¿ add a counter for while loop to exit after 3 
                                        WhileCount = WhileCount + 1
                                        for counter in range(0,105-len(Columns)):
                                                id="wnd[1]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/tblSAPLCNFATC_ALLE_FELDER/txtALLE_FELDER-SCRTEXT[0,0]"
                                                col=self.session.findById(id).text
                                                
                                                if col in Columns:
                                                        self.session.findById("wnd[1]/usr/ssubRAHMEN:SAPLCNFA:0111/subAUSWAHL:SAPLCNFA:0140/btnAUSWAEHLEN").press()
                                                        selected_cols_coun=selected_cols_coun+1
                                                        if selected_cols_coun == len(Columns): break
                                                else:
                                                        self.session.findById("wnd[1]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/tblSAPLCNFATC_ALLE_FELDER").verticalScrollbar.position = counter+1
                                                if selected_cols_coun == len(Columns) or WhileCount == 3:
                                                        WhileCondition = False
                                                        break
                                                
                                self.session.findById("wnd[1]/tbar[0]/btn[0]").press()

                        time.sleep(2)
                        #Set Full characteres len for all fields
                        self.session.findById("wnd[0]/tbar[1]/btn[5]").press()
                        self.session.findById("wnd[1]/usr/ssubRAHMEN:SAPLCNFA:0111/subAKT_FELDER:SAPLCNFA:0120/tblSAPLCNFATC_AKT_FELDER/txtAKT_FELDER-SCRTEXT[0,0]").setFocus()
                        self.session.findById("wnd[1]/usr/ssubRAHMEN:SAPLCNFA:0111/subAKT_FELDER:SAPLCNFA:0120/tblSAPLCNFATC_AKT_FELDER/txtAKT_FELDER-SCRTEXT[0,0]").caretPosition = 11
                        time.sleep(1)
                        self.session.findById("wnd[1]/tbar[0]/btn[6]").press()
                        time.sleep(1)
                        self.session.findById("wnd[1]/usr/ssubRAHMEN:SAPLCNFA:0111/subANZAHL:SAPLCNFA:0150/subANZAHL:SAPLCNFA:0151/btnORIGINALLAENGE_HOLEN").press()
                        time.sleep(1)
                        self.session.findById("wnd[1]/tbar[0]/btn[7]").press()
                        time.sleep(1)
                        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                        time.sleep(1)
                        self.session.findById("wnd[0]/mbar/menu[0]/menu[11]/menu[6]").Select()
                        #self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select() <---text with taps
                        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                        self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = ouput_Path
                        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = filename
                        self.session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
                        self.session.findById("wnd[1]/tbar[0]/btn[11]").press()
                        time.sleep(1)
                        self.session.findById("wnd[0]/tbar[0]/btn[15]").press()
                        self.session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
                        time.sleep(3)
                else:
                        print("Please Add values to extract")
                return response
        #----------------------------------------------------------------------------------CJI3----------------------------------------------------------------------------------------
        def get_CJI3(self, 
                PD=[""],
                Sales_document=[""],
                WBS_element=[""],
                Network_order = [""],
                Activity=[""],
                Materials_in_network=[""],
                Cost_elemet=[""],
                Posting_date=["",""],
                Layout="",
                filename="",
                ouput_Path=config_path
                ):
                if PD[0]!="" or Sales_document[0]!="" or WBS_element[0]!="" or Network_order[0]!="" or Activity[0]!="" or Materials_in_network[0]!="":
                        #Seting variables values
                        if len(filename)<1:
                                default_filename="CJI3_" + PD[0] +"_"+ Sales_document[0] +"_"+ WBS_element[0] +"_"+ Network_order[0] +"_"+ Activity[0] +"_"+ Materials_in_network[0] +".txt"
                                default_filename = default_filename.replace("*","")
                                default_filename = default_filename.replace("__","_")
                                default_filename = default_filename.replace("__","_")
                                default_filename = default_filename.replace("__","_")
                                default_filename = default_filename.replace("_.txt",".txt")
                                filename=default_filename
                        else:
                                if filename.find(".") >= 0:
                                        filename=filename[0:filename.find(".")+1]+".txt"
                                else:
                                        filename=filename + ".txt"

                        if ouput_Path[-1]!="/":
                                ouput_Path = ouput_Path+"/"
                        file_exists = exists(ouput_Path + filename )
                        if file_exists :
                                remove(ouput_Path + filename)
                        
                        # if Posting_date[0]=="":
                        #         Posting_date[0]=date.today().strftime("%m-%d-%Y")
                        #         Posting_date[1]=date.today().strftime("%m-%d-%Y")
                        #open transaction
                        self.session.findById("wnd[0]/tbar[0]/okcd").text = "CJI3"
                        self.session.findById("wnd[0]").sendVKey (0)
                        time.sleep(1)
                        if self.extract_counter==0: #Set Controlling Area
                                #choose Act type
                                self.session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").text = "1000"
                                self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                                time.sleep(2)
                        self.extract_counter=self.extract_counter+1
                        
                        #select variant
                        if self.session.Children.Count > 1 :
                                self.session.findById("wnd[1]/usr/ctxtTCNT-PROF_DB").text = "Z00000000001"
                                self.session.findById("wnd[1]").sendVKey (0)
                                time.sleep(2)

                        #fill fields in transaction
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_CN_PROJN_%_APP_%-VALU_PUSH"  ,value_list= PD ,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id= "wnd[0]/usr/btn%_CN_VBELN_%_APP_%-VALU_PUSH" ,value_list= Sales_document ,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_CN_PSPNR_%_APP_%-VALU_PUSH"  ,value_list=  WBS_element,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_CN_NETNR_%_APP_%-VALU_PUSH"  ,value_list= Network_order ,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id= "wnd[0]/usr/btn%_CN_ACTVT_%_APP_%-VALU_PUSH" ,value_list= Activity ,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id= "wnd[0]/usr/btn%_CN_MATNR_%_APP_%-VALU_PUSH" ,value_list= Materials_in_network ,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id= "wnd[0]/usr/btn%_R_KSTAR_%_APP_%-VALU_PUSH" ,value_list= Cost_elemet ,ClipBoard_path=ouput_Path)                 

                        #---Set Posting Date
                        self.session.findById("wnd[0]/usr/ctxtR_BUDAT-LOW").text = Posting_date[0]
                        self.session.findById("wnd[0]/usr/ctxtR_BUDAT-HIGH").text = Posting_date[1]

                        #---Set Layout
                        self.session.findById("wnd[0]/usr/ctxtP_DISVAR").text = Layout

                        #--Set Hits for max memory
                        self.session.findById("wnd[0]/usr/btnBUT1").press()
                        self.session.findById("wnd[1]/usr/txtKAEP_SETT-MAXSEL").text = "999999999"
                        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()

                        #RUN
                        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()  # Run  <------------------------------------RUN -------------------
                        time.sleep(1)
                        # <----------------------------------Validate transaction running----------------------------------------------------------
                        
                        response=self.validate_run_response()
                        if not response: return response
                        time.sleep(3)

                        self.session.findById("wnd[0]/tbar[1]/btn[45]").press()
                        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                        self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = ouput_Path
                        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = filename
                        self.session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
                        self.session.findById("wnd[1]/tbar[0]/btn[11]").press()
                        
                        self.session.findById("wnd[0]/tbar[0]/btn[15]").press()
                        self.session.findById("wnd[0]/tbar[0]/btn[15]").press()
                        time.sleep(3)
                else:
                        print("Please Add values to extract")
                return response
        #----------------------------------------------------------------------------------CNji5---------------------------------------------------------------------------------------
        def get_CJI5(self, 
                PD=[""],
                Sales_document=[""],
                WBS_element=[""],
                Network_order = [""],
                Activity=[""],
                Materials_in_network=[""],
                Cost_elemet=[""],
                Posting_date=["",""],
                Layout="",
                filename="",
                ouput_Path=config_path
                ):
                if PD[0]!="" or Sales_document[0]!="" or WBS_element[0]!="" or Network_order[0]!="" or Activity[0]!="" or Materials_in_network[0]!="":
                        #Seting variables values
                        if len(filename)<1:
                                default_filename="CJI5_" + PD[0] +"_"+ Sales_document[0] +"_"+ WBS_element[0] +"_"+ Network_order[0] +"_"+ Activity[0] +"_"+ Materials_in_network[0] +".txt"
                                default_filename = default_filename.replace("*","")
                                default_filename = default_filename.replace("__","_")
                                default_filename = default_filename.replace("__","_")
                                default_filename = default_filename.replace("__","_")
                                default_filename = default_filename.replace("_.txt",".txt")
                                filename=default_filename
                        else:
                                if filename.find(".") >= 0:
                                        filename=filename[0:filename.find(".")+1]+".txt"
                                else:
                                        filename=filename + ".txt"

                        if ouput_Path[-1]!="/":
                                ouput_Path = ouput_Path+"/"
                        file_exists = exists(ouput_Path + filename )
                        if file_exists :
                                remove(ouput_Path + filename)
                        
                        # if Posting_date[0]=="":
                        #         Posting_date[0]=date.today().strftime("%m-%d-%Y")
                        #         Posting_date[1]=date.today().strftime("%m-%d-%Y")
                        #open transaction
                        self.session.findById("wnd[0]/tbar[0]/okcd").text = r"/NCJI5"
                        self.session.findById("wnd[0]").sendVKey (0)
                        time.sleep(1)

                        if self.extract_counter==0:
                                #choose Act type
                                self.session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").text = "1000"
                                self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                                time.sleep(1)
                        self.extract_counter=self.extract_counter+1

                        #select variant
                        if self.session.Children.Count > 1 :
                                self.session.findById("wnd[1]/usr/ctxtTCNT-PROF_DB").text = "Z00000000001"
                                self.session.findById("wnd[1]").sendVKey (0)
                                time.sleep(2)

                        #fill fields in transaction
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_CN_PROJN_%_APP_%-VALU_PUSH"  ,value_list= PD ,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id= "wnd[0]/usr/btn%_CN_VBELN_%_APP_%-VALU_PUSH" ,value_list= Sales_document ,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_CN_PSPNR_%_APP_%-VALU_PUSH"  ,value_list=  WBS_element,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_CN_NETNR_%_APP_%-VALU_PUSH"  ,value_list= Network_order ,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id= "wnd[0]/usr/btn%_CN_ACTVT_%_APP_%-VALU_PUSH" ,value_list= Activity ,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id= "wnd[0]/usr/btn%_CN_MATNR_%_APP_%-VALU_PUSH" ,value_list= Materials_in_network ,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id= "wnd[0]/usr/btn%_R_KSTAR_%_APP_%-VALU_PUSH" ,value_list= Cost_elemet ,ClipBoard_path=ouput_Path)

                        #---Set Posting Date

                        self.session.findById("wnd[0]/usr/ctxtR_OBDAT-LOW").text = Posting_date[0]
                        self.session.findById("wnd[0]/usr/ctxtR_OBDAT-HIGH").text = Posting_date[1]

                        #---Set Layout
                        self.session.findById("wnd[0]/usr/ctxtP_DISVAR").text = Layout
                        time.sleep(1)

                        #--Set Hits for max memory
                        self.session.findById("wnd[0]/usr/btnBUT1").press()
                        self.session.findById("wnd[1]/usr/txtKAEP_SETT-MAXSEL").text = "999999999"
                        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()  # Run  <------------------------------------RUN
                        time.sleep(1)

                        # <----------------------------------Validate transaction running----------------------------------------------------------
                        response=self.validate_run_response()
                        if not response: return response
                        time.sleep(3)
                        

                        self.session.findById("wnd[0]/tbar[1]/btn[45]").press()
                        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                        self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = ouput_Path
                        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = filename
                        self.session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
                        self.session.findById("wnd[1]/tbar[0]/btn[11]").press()
                        
                        self.session.findById("wnd[0]/tbar[0]/btn[15]").press()
                        self.session.findById("wnd[0]/tbar[0]/btn[15]").press()
                        time.sleep(3)
                else:
                        print("Please Add values to extract")
                return response
        #----------------------------------------------------------------------------------zsnap---------------------------------------------------------------------------------------
        def get_ZSNAP(self,
                customer_number=0,
                customer_PO_number=[""],
                Date_received=["",""],
                Changed_time=[""],
                Date_last_855=[""],
                Material = [""],
                Buyer_name=[""],
                Circular_PO_type=[""],
                rejection_reason=[""],
                Posting_date=["",""],
                Layout="", 
                filename="",
                ouput_Path=config_path
                ):

                if customer_number!=0 or customer_PO_number[0]!="" or Date_received[0]!="" or Changed_time[0]!="" or Date_last_855[0]!="" or Material[0]!="" or Buyer_name[0]!="" or Circular_PO_type[0]!="":
                        #Seting variables values
                        if len(filename)<1:
                                default_filename="ZSNAP_" + str(customer_number) + customer_PO_number[0] +"_"+ Date_received[0] +"_"+ Changed_time[0] +"_"+ Date_last_855[0] +"_"+ Material[0] +"_"+ Buyer_name[0] +".txt"
                                default_filename = default_filename.replace("*","")
                                default_filename = default_filename.replace("__","_")
                                default_filename = default_filename.replace("__","_")
                                default_filename = default_filename.replace("__","_")
                                default_filename = default_filename.replace("_.txt",".txt")
                                filename=default_filename
                        else:
                                if filename.find(".") >= 0:
                                        filename=filename[0:filename.find(".")+1]+".txt"
                                else:
                                        filename=filename + ".txt"
                                        
                        if ouput_Path[-1]!="/":
                                ouput_Path = ouput_Path+"/"
                        file_exists = exists(ouput_Path + filename )
                        if file_exists :
                                remove(ouput_Path + filename)
                        
                        # if Posting_date[0]=="":
                        #         Posting_date[0]=date.today().strftime("%m-%d-%Y")
                        #         Posting_date[1]=date.today().strftime("%m-%d-%Y")


                        #open transaction----------------------------------------------------------------------------------------------------------
                        self.session.findById("wnd[0]/tbar[0]/okcd").text = "ZSNAP_D"
                        self.session.findById("wnd[0]/tbar[0]/btn[0]").press()
                        time.sleep(1)
                        #set transaction parameters
                        self.session.findById("wnd[0]/usr/ctxtP_KUNNR").text  = customer_number
                        self.session.findById("wnd[0]/usr/radRB_PRI").Select()
                        
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_BSTKD_%_APP_%-VALU_PUSH",value_list=customer_PO_number,ClipBoard_path=ouput_Path)

                        if len(Date_received) == 2: 
                                self.session.findById("wnd[0]/usr/ctxtS_ERDAT-LOW").text = Date_received[0]
                                self.session.findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").text = Date_received[1]
                        else:
                                self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_ERDAT_%_APP_%-VALU_PUSH",value_list=Date_received,ClipBoard_path=ouput_Path)

                        if len(Changed_time) == 2: 
                                self.session.findById("wnd[0]/usr/ctxtS_UZEIT-LOW").text = Changed_time[0]
                                self.session.findById("wnd[0]/usr/ctxtS_UZEIT-HIGH").text = Changed_time[1]
                        else:
                                self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_UZEIT_%_APP_%-VALU_PUSH",value_list=Changed_time,ClipBoard_path=ouput_Path)
                        
                        if len(Date_last_855) == 2: 
                                self.session.findById("wnd[0]/usr/ctxtS_855-LOW").text = Date_last_855[0]
                                self.session.findById("wnd[0]/usr/ctxtS_855-HIGH").text = Date_last_855[1]
                        else:
                                self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_855_%_APP_%-VALU_PUSH",value_list=Date_last_855,ClipBoard_path=ouput_Path)

                        
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH",value_list=Material,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_BUYER_%_APP_%-VALU_PUSH",value_list=Buyer_name,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_BSARK_%_APP_%-VALU_PUSH",value_list=Circular_PO_type,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_ABGRU_%_APP_%-VALU_PUSH",value_list=rejection_reason,ClipBoard_path=ouput_Path)
                        
                        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()  # Run  <------------------------------------RUN
                        time.sleep(2)

                        # <----------------------------------Validate transaction running----------------------------------------------------------
                        response=self.validate_run_response()
                        if not response: return response
                        time.sleep(3)
                        

                        #setlayout
                        self.session.findById("wnd[0]/tbar[1]/btn[33]").press()
                        self.session.findById("wnd[1]/usr/lbl[1,3]").setFocus()
                        self.session.findById("wnd[1]/usr/lbl[1,3]").caretPosition = 0
                        self.session.findById("wnd[1]").sendVKey (29)
                        self.session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = Layout
                        self.session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 10
                        self.session.findById("wnd[2]/tbar[0]/btn[0]").press()
                        self.session.findById("wnd[1]/usr/lbl[1,3]").setFocus()
                        #self.session.findById("wnd[1]/usr/lbl[1,3]").caretPosition = 6
                        self.session.findById("wnd[1]").sendVKey (2)


                        self.session.findById("wnd[0]/tbar[1]/btn[45]").press()
                        #self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select()
                        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                        self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = ouput_Path
                        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = filename
                        self.session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
                        self.session.findById("wnd[1]/tbar[0]/btn[11]").press()
                        
                        #exit of transaction
                        self.session.findById("wnd[0]/tbar[0]/btn[15]").press()
                        self.session.findById("wnd[0]/tbar[0]/btn[15]").press()
                        time.sleep(3)
                else:
                        print("Please Add values to extract")
                return response
        #----------------------------------------------------------------------------------zzcustmon---------------------------------------------------------------------------------------
        def get_Zzcustmon(self,
                sales_org="1263", 
                sales_office=[""],
                Sales_group=[""],
                Order_type=[""],
                Sales_order_number = [""],
                Site_id=[""],
                Network_number=[""],
                WBS_element=[""],
                Order_Reason=[""],
                Reason_for_rejection_code=[""],
                Created_by=[""],
                Collective_number=[""],
                Created_on=[""],
                Customer_req_date=[""],
                Our_confirmed_date=[""],
                Sold_to_party=[""],
                Output_Type=[""],
                Plant=[""],
                Item_Category=[""],
                Planned_GI_date=["",""],
                Actual_GI_date=["",""],
                Planned_delivery_date=["",""],
                Actual_delivery_date=["",""],
                SO_line_create_date=["",""],
                POD_status=[""],
                Material=[""],
                Customer_PO=[""],
                Contract_No=[""],
                Billing_Plan_ref_nr=[""],
                Verdi_Site_Name=[""],
                Material_Avail_date=[""],
                Final_External_Customer=[""],
                End_user_for_Trade=[""],
                Delivery_manager=[""],
                ref_SO_for_CONS=[""],
                Supply_hub_plan=[""],
                Layout=r"",
                filename="",
                ouput_Path=config_path
                ):
                if Sales_order_number[0]!="" or Site_id[0]!="" or Network_number[0]!="" or WBS_element[0]!="" or Customer_PO[0]!="":
                        #Seting variables values
                        if len(filename)<1:
                                default_filename="ZZCUSTMON_" + Sales_order_number[0] +"_"+ Site_id[0] +"_"+ WBS_element[0] +"_"+ Network_number[0] +"_"+ WBS_element[0] +"_"+ Customer_PO[0] +".txt"
                                default_filename = default_filename.replace("*","")
                                default_filename = default_filename.replace("__","_")
                                default_filename = default_filename.replace("__","_")
                                default_filename = default_filename.replace("__","_")
                                default_filename = default_filename.replace("__","_")
                                default_filename = default_filename.replace("_.txt",".txt")
                                filename=default_filename
                        else:
                                if filename.find(".") >= 0:
                                        filename=filename[0:filename.find(".")+1]+".txt"
                                else:
                                        filename=filename + ".txt"
                                        
                        if ouput_Path[-1]!="/":
                                ouput_Path = ouput_Path+"/"
                        file_exists = exists(ouput_Path + filename )
                        if file_exists :
                                remove(ouput_Path + filename)
                        
                        #open transaction
                        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nZZCUSTMON"
                        self.session.findById("wnd[0]").sendVKey(0)
                        self.session.findById("wnd[0]/usr/ctxtP_VKORG").text = sales_org
                        self.session.findById("wnd[0]/usr/chkP_CHKBOX").selected = True
                        time.sleep(1)

                        # #choose Act type
                        # self.session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").text = "1000"
                        # self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                        # time.sleep(1)
                        # #select variant
                        # if self.session.Children.Count > 1 :
                        #         self.session.findById("wnd[1]/usr/ctxtTCNTT-PROFID").text = "Z00000000001"
                        #         self.session.findById("wnd[1]").sendVKey (0)


                        #fill fields in transaction

                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_VKBUR_%_APP_%-VALU_PUSH"    ,value_list= sales_office ,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_VKGRP_%_APP_%-VALU_PUSH"    ,value_list=Sales_group  ,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_AUART_%_APP_%-VALU_PUSH"    ,value_list=Order_type  ,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_VBELN_%_APP_%-VALU_PUSH"    ,value_list=Sales_order_number  ,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_SITID_%_APP_%-VALU_PUSH"    ,value_list=Site_id  ,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_NPLNR_%_APP_%-VALU_PUSH"    ,value_list=Network_number  ,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_PROJN_%_APP_%-VALU_PUSH"    ,value_list=WBS_element  ,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_AUGRU_%_APP_%-VALU_PUSH"    ,value_list=Order_Reason  ,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_ABGRU_%_APP_%-VALU_PUSH"    ,value_list=Reason_for_rejection_code  ,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_ERNAM_%_APP_%-VALU_PUSH"    ,value_list=Created_by  ,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_SUBMI_%_APP_%-VALU_PUSH"    ,value_list=Collective_number  ,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_ERDAT_%_APP_%-VALU_PUSH"    ,value_list=Created_on  ,ClipBoard_path=ouput_Path)

                        if len(Customer_req_date) == 2: 
                                self.session.findById("wnd[0]/usr/ctxtS_VDATU-LOW").text = Customer_req_date[0]
                                self.session.findById("wnd[0]/usr/ctxtS_VDATU-HIGH").text = Customer_req_date[1]
                        else:
                                self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_VDATU_%_APP_%-VALU_PUSH"    ,value_list=Customer_req_date  ,ClipBoard_path=ouput_Path)

                        if len(Our_confirmed_date) == 2: 
                                self.session.findById("wnd[0]/usr/ctxtS_EDATU-LOW").text = Our_confirmed_date[0]
                                self.session.findById("wnd[0]/usr/ctxtS_EDATU-HIGH").text = Our_confirmed_date[1]
                        else:
                                self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_EDATU_%_APP_%-VALU_PUSH"    ,value_list=Our_confirmed_date  ,ClipBoard_path=ouput_Path)                       
                        
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_KUNNR_%_APP_%-VALU_PUSH"    ,value_list=Sold_to_party  ,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_KSCHL_%_APP_%-VALU_PUSH"    ,value_list=Output_Type ,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_WERKS_%_APP_%-VALU_PUSH"    ,value_list=Plant  ,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_PSTYV_%_APP_%-VALU_PUSH"    ,value_list=Item_Category  ,ClipBoard_path=ouput_Path)

                        if Planned_GI_date[0] != "": 
                                self.session.findById("wnd[0]/usr/ctxtS_WADAT-LOW").text = Planned_GI_date[0]
                                self.session.findById("wnd[0]/usr/ctxtS_WADAT-HIGH").text = Planned_GI_date[1]
                        else:
                                self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_WADAT_%_APP_%-VALU_PUSH"    ,value_list=Planned_GI_date  ,ClipBoard_path=ouput_Path)

                        if Actual_GI_date[0] != "": 
                                self.session.findById("wnd[0]/usr/ctxtS_WADATI-LOW").text = Actual_GI_date[0]
                                self.session.findById("wnd[0]/usr/ctxtS_WADATI-HIGH").text = Actual_GI_date[1]
                        else:
                                self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_WADATI_%_APP_%-VALU_PUSH"   ,value_list=Actual_GI_date  ,ClipBoard_path=ouput_Path)

                        if Planned_delivery_date[0] != "": 
                                self.session.findById("wnd[0]/usr/ctxtS_LFDAT-LOW").text = Planned_delivery_date[0]
                                self.session.findById("wnd[0]/usr/ctxtS_LFDAT-HIGH").text = Planned_delivery_date[1]
                        else:
                                self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_LFDAT_%_APP_%-VALU_PUSH"    ,value_list=Planned_delivery_date  ,ClipBoard_path=ouput_Path)

                        if Actual_delivery_date[0] != "": 
                                self.session.findById("wnd[0]/usr/ctxtS_PODAT-LOW").text = Actual_delivery_date[0]
                                self.session.findById("wnd[0]/usr/ctxtS_PODAT-HIGH").text = Actual_delivery_date[1]
                        else:
                                self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_PODAT_%_APP_%-VALU_PUSH"    ,value_list=Actual_delivery_date  ,ClipBoard_path=ouput_Path)

                        if SO_line_create_date[0] != "": 
                                self.session.findById("wnd[0]/usr/ctxtS_ERDATI-LOW").text = SO_line_create_date[0]
                                self.session.findById("wnd[0]/usr/ctxtS_ERDATI-HIGH").text = SO_line_create_date[1]
                        else:
                                self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_ERDATI_%_APP_%-VALU_PUSH"   ,value_list=SO_line_create_date  ,ClipBoard_path=ouput_Path)

                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_PDSTA_%_APP_%-VALU_PUSH"    ,value_list=POD_status  ,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH"    ,value_list= Material ,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_BSTNK_%_APP_%-VALU_PUSH"    ,value_list=Customer_PO  ,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_ZZCON_%_APP_%-VALU_PUSH"    ,value_list= Contract_No ,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_FPLNR_%_APP_%-VALU_PUSH"    ,value_list=Billing_Plan_ref_nr  ,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_SITNM_%_APP_%-VALU_PUSH"    ,value_list=Verdi_Site_Name  ,ClipBoard_path=ouput_Path)

                        if Material_Avail_date[0] != "":  
                                self.session.findById("wnd[0]/usr/ctxtS_MBDAT-LOW").text = Material_Avail_date[0]
                                self.session.findById("wnd[0]/usr/ctxtS_MBDAT-HIGH").text = Material_Avail_date[1]
                        else:
                                self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_MBDAT_%_APP_%-VALU_PUSH"   ,value_list=Material_Avail_date  ,ClipBoard_path=ouput_Path)

                        
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_KUNNR1_%_APP_%-VALU_PUSH"   ,value_list=Final_External_Customer  ,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_KUNNR2_%_APP_%-VALU_PUSH"   ,value_list=End_user_for_Trade  ,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_ORDRES_%_APP_%-VALU_PUSH"   ,value_list=Delivery_manager  ,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_CONSID_%_APP_%-VALU_PUSH"   ,value_list=ref_SO_for_CONS  ,ClipBoard_path=ouput_Path)

                        # #---Set Posting Date
                        # self.session.findById("wnd[0]/usr/ctxtR_OBDAT-LOW").text = Posting_date[0]
                        # self.session.findById("wnd[0]/usr/ctxtR_OBDAT-HIGH").text = Posting_date[1]

                        #---Set Layout
                        self.session.findById("wnd[0]/usr/ctxtP_VARI").text = Layout
                        self.session.findById("wnd[0]").sendVKey(0)
                        time.sleep(1)
                        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()  # Run  <------------------------------------RUN
                        time.sleep(2)

                        # <----------------------------------Validate transaction running----------------------------------------------------------
                        response=self.validate_run_response()
                        if not response: return response
                        time.sleep(3)
                        
                                        
                        self.session.findById("wnd[0]/tbar[1]/btn[45]").press()
                        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                        self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = ouput_Path
                        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = filename
                        self.session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
                        self.session.findById("wnd[1]/tbar[0]/btn[11]").press()
                        
                        self.session.findById("wnd[0]/tbar[0]/btn[15]").press()
                        #self.session.findById("wnd[0]/tbar[0]/btn[15]").press()
                        time.sleep(3)
                else:
                        print("Please Add values to extract")
                return response
        #----------------------------------------------------------------------------------zcheckqtc---------------------------------------------------------------------------------------
        def get_zcheckqtc(self,
                sales_org="1263",
                Distribution_channel=[""], 
                Sales_order_type=[""],
                Sales_office=[""],
                Customer_id=[""],
                Sales_order = [""],
                Sales_order_date = [""],
                Profit_centre=[""],
                P_code=[""],
                Currency_code=[""],
                Business_unit=[""],
                Ericsson_contract_number=[""],
                Project=[""],
                WBS_element=[""],
                Site_id=[""],
                Layout="",
                restrict_OP=False,
                blocked_docs=True,
                ignore_comp=False,
                date_to="",
                filename="",
                ouput_Path=config_path
                ):
                if Distribution_channel[0]!="" or Site_id[0]!="" or Project[0]!="" or WBS_element[0]!="" or Sales_order_type[0]!="" or Sales_office[0]!="" or Customer_id[0]!="" or Sales_order[0]!="" or Sales_order_date[0]!="":
                        #Seting variables values
                        if len(filename)<1:
                                default_filename="ZCHECKQTC_" + Project[0] +"_"+ Site_id[0] +"_"+ WBS_element[0] +"_"+ Customer_id[0] +"_"+ Sales_order[0] +"_"+ Sales_order_date[0] +".txt"
                                default_filename = default_filename.replace("*","")
                                default_filename = default_filename.replace("__","_")
                                default_filename = default_filename.replace("__","_")
                                default_filename = default_filename.replace("__","_")
                                default_filename = default_filename.replace("__","_")
                                default_filename = default_filename.replace("_.txt",".txt")
                                filename=default_filename
                        else:
                                if filename.find(".") >= 0:
                                        filename=filename[0:filename.find(".")+1]+".txt"
                                else:
                                        filename=filename + ".txt"
                                        
                        if ouput_Path[-1]!="/":
                                ouput_Path = ouput_Path+"/"
                        file_exists = exists(ouput_Path + filename )
                        if file_exists :
                                remove(ouput_Path + filename)
                        
                        #if date_to=="":
                        #        date_to=date.today().strftime("%m-%d-%Y")

                        #open transaction
                        self.session.findById("wnd[0]/tbar[0]/okcd").text = "zzcheckqtc"
                        self.session.findById("wnd[0]").sendVKey(0)
                        self.session.findById("wnd[0]/usr/ctxtP_VKORG").text = sales_org
                        time.sleep(1)

                        #fill fields in transaction
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_VTWEG_%_APP_%-VALU_PUSH"    ,value_list= Distribution_channel ,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_AUART_%_APP_%-VALU_PUSH"    ,value_list=Sales_order_type  ,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_VKBUR_%_APP_%-VALU_PUSH"    ,value_list=Sales_office  ,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_KUNNR_%_APP_%-VALU_PUSH"    ,value_list=Customer_id  ,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_VBELN_%_APP_%-VALU_PUSH"    ,value_list=Sales_order  ,ClipBoard_path=ouput_Path)
                        if len(Sales_order_date) == 2: 
                                self.session.findById("wnd[0]/usr/ctxtS_AUDAT-LOW").text = Sales_order_date[0]
                                self.session.findById("wnd[0]/usr/ctxtS_AUDAT-HIGH").text = Sales_order_date[1]
                        else:
                                self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_AUDAT_%_APP_%-VALU_PUSH"    ,value_list=Sales_order_date  ,ClipBoard_path=ouput_Path)

                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_PRCTR_%_APP_%-VALU_PUSH"    ,value_list=Profit_centre  ,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_MVGR1_%_APP_%-VALU_PUSH"    ,value_list=P_code  ,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_WAERK_%_APP_%-VALU_PUSH"    ,value_list=Currency_code  ,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_KVGR1_%_APP_%-VALU_PUSH"    ,value_list=Business_unit  ,ClipBoard_path=ouput_Path)

                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_ZZCON_%_APP_%-VALU_PUSH"    ,value_list=Ericsson_contract_number  ,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_PSPID_%_APP_%-VALU_PUSH"    ,value_list=Project  ,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_POSID_%_APP_%-VALU_PUSH"    ,value_list=WBS_element  ,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_SITEID_%_APP_%-VALU_PUSH"    ,value_list=Site_id ,ClipBoard_path=ouput_Path)


                        self.session.findById("wnd[0]/usr/chkP_SYSDT").selected = restrict_OP
                        self.session.findById("wnd[0]/usr/chkP_BLKCHK").selected = blocked_docs
                        self.session.findById("wnd[0]/usr/chkP_IGNBIL").selected = ignore_comp
                        self.session.findById("wnd[0]/usr/ctxtP_BILLDT").text = date_to
                        self.session.findById("wnd[0]/usr/chkP_IGNBIL").setFocus
                        # #---Set Posting Datewnd
                        # self.session.findById("wnd[0]/usr/ctxtR_OBDAT-LOW").text = Posting_date[0]
                        # self.session.findById("wnd[0]/usr/ctxtR_OBDAT-HIGH").text = Posting_date[1]

                        # #---Set Layout
                        # self.session.findById("wnd[0]/usr/ctxtP_VARI").text = Layout
                        # self.session.findById("wnd[0]").sendVKey(0)
                        # time.sleep(1)
                        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()  # Run  <------------------------------------RUN
                        time.sleep(2)

                        #Validate transaction
                        response=self.validate_run_response()
                        if not response: return response
                        time.sleep(3)

                        #Search and apply Layout
                        #--------------------------------------------------SET columns---------------------------------------------------
                        self.session.findById("wnd[0]/tbar[1]/btn[33]").press()
                        time.sleep(2)
                        
                        self.session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").contextMenu()
                        self.session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectContextMenuItem ("&FILTER")
                        self.session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = Layout
                        self.session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 11
                        self.session.findById("wnd[2]/tbar[0]/btn[0]").press()
                        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                                                        
                        self.session.findById("wnd[0]/tbar[1]/btn[45]").press()
                        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                        self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = ouput_Path
                        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = filename
                        self.session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
                        self.session.findById("wnd[1]/tbar[0]/btn[11]").press()
                        
                        self.session.findById("wnd[0]/tbar[0]/btn[15]").press()
                        self.session.findById("wnd[0]/tbar[0]/btn[15]").press()
                        time.sleep(3)
                else:
                        print("Please Add values to extract")
                return response
        #-----------------------------------------------------------------------------------CNS46---------------------------------------------------------------------------------------
        def get_CN46N(self, 
                PD=[""],
                Sales_document=[""],
                WBS_element=[""],
                Network_order = [""],
                layout="",
                level=[1,99],
                filename="",
                ouput_Path=config_path,
                Columns=[""]
                ):
                if PD[0]!="" or Sales_document[0]!="" or WBS_element[0]!="" or Network_order[0]!="" :

                        if len(filename)<1:
                                default_filename="CN46N_" + PD[0] +"_"+ Sales_document[0] +"_"+ WBS_element[0] +"_"+ Network_order[0] +"_" +".txt"
                                default_filename = default_filename.replace("*","")
                                default_filename = default_filename.replace("__","_")
                                default_filename = default_filename.replace("__","_")
                                default_filename = default_filename.replace("__","_")
                                default_filename = default_filename.replace("_.txt",".txt")
                                filename=default_filename
                        else:
                                if filename.find(".") >= 0:
                                        filename=filename[0:filename.find(".")+1]+".txt"
                                else:
                                        filename=filename + ".txt"

                        if ouput_Path[-1]!="/":
                                ouput_Path = ouput_Path+"/"

                        file_exists = exists(ouput_Path + filename )
                        if file_exists :
                                remove(ouput_Path + filename)
                        #open transaction
                        self.session.findById("wnd[0]").maximize
                        self.session.findById("wnd[0]/tbar[0]/okcd").text = "CN46N"
                        self.session.findById("wnd[0]").sendVKey(0)

                        if self.session.Children.Count > 1 :
                                self.session.findById("wnd[1]/usr/ctxtTCNT-PROF_DB").text = "Z00000000001"
                                self.session.findById("wnd[1]").sendVKey (0)

                        #fill fields in transaction
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_CN_PROJN_%_APP_%-VALU_PUSH",value_list=PD,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_CN_VBELN_%_APP_%-VALU_PUSH",value_list=Sales_document,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_CN_PSPNR_%_APP_%-VALU_PUSH",value_list=WBS_element,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_CN_NETNR_%_APP_%-VALU_PUSH",value_list=Network_order,ClipBoard_path=ouput_Path)
                        #self.set_transaction_parameter(id="wnd[0]/usr/btn%_CN_ACTVT_%_APP_%-VALU_PUSH",value_list=Activity,ClipBoard_path=ouput_Path)
                        #self.set_transaction_parameter(id="wnd[0]/usr/btn%_CN_MATNR_%_APP_%-VALU_PUSH",value_list=Materials_in_network,ClipBoard_path=ouput_Path)
                        
                        #set level for extraction
                        self.session.findById("wnd[0]/usr/txtCN_STUFE-LOW").text = min(level)
                        self.session.findById("wnd[0]/usr/txtCN_STUFE-HIGH").text = max(level)
                        self.session.findById("wnd[0]").sendVKey (0)

                        #Layout
                        self.session.findById("wnd[0]/usr/ctxtP_DISVAR").text = layout          
                        
                        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()  # Run  <--------------------
                        time.sleep(1)
                        # ----------------------------------Validate transaction running ------------------------------------------------------
                        response=self.validate_run_response()
                        if not response: return response
                        time.sleep(3)


                        #--------------------------------------------------SET columns---------------------------------------------------
                        if Columns[0] != "" :   
                                # self.session.findById("wnd[0]/tbar[1]/btn[5]").press()
                                # self.session.findById("wnd[1]/usr/ssubRAHMEN:SAPLCNFA:0111/subAUSWAHL:SAPLCNFA:0140/btnALLE_NICHT_AUSWAEHLEN").press()
                                self.session.findById("wnd[1]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/tblSAPLCNFATC_ALLE_FELDER").verticalScrollbar.position = 0
                                self.session.findById("wnd[0]/usr/cntlALVCONTAINER/shellcont/shell").pressToolbarContextButton ("&MB_VARIANT")
                                self.session.findById("wnd[0]/usr/cntlALVCONTAINER/shellcont/shell").selectContextMenuItem ("&COL0")

                                selected_cols_coun=0
                                for counter in range(0,105-len(Columns)):
                                        id="wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell"
                                        col=self.session.findById(id).text
                                        
                                        if col in Columns:
                                                self.session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press()
                                                selected_cols_coun=selected_cols_coun+1
                                                if selected_cols_coun == len(Columns): break
                                        else:
                                                #self.session.findById("wnd[1]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/tblSAPLCNFATC_ALLE_FELDER").verticalScrollbar.position = counter+1
                                                self.session.findById(id).currentCellRow = counter+1
                                                self.session.findById(id).firstVisibleRow = counter+1
                                                self.session.findById(id).selectedRows = str(counter+1)
                                self.session.findById("wnd[1]/tbar[0]/btn[0]").press()

                        
                        

                        
                        #Set Full characteres len for all fields
                        # self.session.findById("wnd[0]/tbar[1]/btn[5]").press()
                        # self.session.findById("wnd[1]/usr/ssubRAHMEN:SAPLCNFA:0111/subAKT_FELDER:SAPLCNFA:0120/tblSAPLCNFATC_AKT_FELDER/txtAKT_FELDER-SCRTEXT[0,0]").setFocus()
                        # self.session.findById("wnd[1]/usr/ssubRAHMEN:SAPLCNFA:0111/subAKT_FELDER:SAPLCNFA:0120/tblSAPLCNFATC_AKT_FELDER/txtAKT_FELDER-SCRTEXT[0,0]").caretPosition = 11
                        # self.session.findById("wnd[1]/tbar[0]/btn[6]").press()
                        # self.session.findById("wnd[1]/usr/ssubRAHMEN:SAPLCNFA:0111/subANZAHL:SAPLCNFA:0150/subANZAHL:SAPLCNFA:0151/btnORIGINALLAENGE_HOLEN").press()
                        # self.session.findById("wnd[1]/tbar[0]/btn[7]").press()
                        # self.session.findById("wnd[1]/tbar[0]/btn[0]").press()

                        # Export File
                        self.session.findById("wnd[0]/usr/cntlALVCONTAINER/shellcont/shell").pressToolbarContextButton ("&MB_EXPORT")
                        self.session.findById("wnd[0]/usr/cntlALVCONTAINER/shellcont/shell").selectContextMenuItem ("&PC")
                        #self.session.findById("wnd[0]/mbar/menu[0]/menu[11]/menu[6]").Select()
                        #self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select() <---un comment to get text with taps
                        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                        self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = ouput_Path
                        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = filename
                        self.session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
                        self.session.findById("wnd[1]/tbar[0]/btn[11]").press()

                        self.session.findById("wnd[0]/tbar[0]/btn[15]").press()
                        time.sleep(3)
                else:
                        print("Please Add values to extract")

                return response

        #------------------------------------------------------------------------------------ME2J--------------------------------------------------------------------------------------
        def get_ME2J(self, 
                PD=[""],
                Sales_document=[""],
                WBS_element=[""],
                Network_order = [""],
                Activity=[""],
                Layout="/EKAUMIR",
                level=[1,99],
                Purchace_order=True,
                Contract=False,
                Sch_Agmt=True,
                Purchasing_organization=[""],
                Scope_of_list=[""],
                Selection_parameter=[""],
                Document_type=[""],
                Purchasing_group=[""],
                Plant=[""],
                Item_category=[""],
                Account_assignment_category=[""],
                Delivery_date=[""],
                Validity_key_date=[""],
                Range_to=[""],
                Document_no=[""],
                Vendor=[""],
                Delivering_plant=[""],
                Material=[""],
                Material_group=[""],
                Document_date=[""],
                European_article_number=[""],
                Vendor_material_number=[""],
                Vendor_sub_range=[""],
                Action=[""],
                Season=[""],
                Season_year=[""],
                Short_text=[""],
                Vendor_name=[""],
                filename="",
                ouput_Path=config_path,
                Columns=[""]
                ):
                if PD[0]!="" or Sales_document[0]!="" or WBS_element[0]!="" or Network_order[0]!="" :

                        if len(filename)<1:
                                default_filename="ME2J_" + PD[0] +"_"+ Sales_document[0] +"_"+ WBS_element[0] +"_"+ Network_order[0] +"_" +".txt"
                                default_filename = default_filename.replace("*","")
                                default_filename = default_filename.replace("__","_")
                                default_filename = default_filename.replace("__","_")
                                default_filename = default_filename.replace("__","_")
                                default_filename = default_filename.replace("_.txt",".txt")
                                filename=default_filename
                        else:
                                if filename.find(".") >= 0:
                                        filename=filename[0:filename.find(".")+1]+".txt"
                                else:
                                        filename=filename + ".txt"

                        if ouput_Path[-1]!="/":
                                ouput_Path = ouput_Path+"/"

                        file_exists = exists(ouput_Path + filename )
                        if file_exists :
                                remove(ouput_Path + filename)
                        #open transaction
                        self.session.findById("wnd[0]").maximize
                        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/NME2J"
                        self.session.findById("wnd[0]").sendVKey(0)

                        if self.session.Children.Count > 1 :
                                self.session.findById("wnd[1]/usr/ctxtTCNT-PROF_DB").text = "Z00000000001"
                                self.session.findById("wnd[1]").sendVKey (0)

                        #fill fields in transaction
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_CN_PROJN_%_APP_%-VALU_PUSH",value_list=PD,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_CN_VBELN_%_APP_%-VALU_PUSH",value_list=Sales_document,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_CN_PSPNR_%_APP_%-VALU_PUSH",value_list=WBS_element,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_CN_NETNR_%_APP_%-VALU_PUSH",value_list=Network_order,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_CN_ACTVT_%_APP_%-VALU_PUSH",value_list=Activity,ClipBoard_path=ouput_Path)

                        self.session.findById("wnd[0]/usr/chkEK_SELKB").selected = Purchace_order
                        self.session.findById("wnd[0]/usr/chkEK_SELKK").selected = Contract
                        self.session.findById("wnd[0]/usr/chkEK_SELKL").selected = Sch_Agmt

                        #set level for extraction
                        self.session.findById("wnd[0]/usr/txtCN_STUFE-LOW").text = min(level)
                        self.session.findById("wnd[0]/usr/txtCN_STUFE-HIGH").text = max(level)
                        self.session.findById("wnd[0]").sendVKey (0)

                        self.session.findById("wnd[0]/usr/chkEK_SELKK").setFocus

                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_EKORG_%_APP_%-VALU_PUSH",value_list=Purchasing_organization,ClipBoard_path=ouput_Path)
                        self.session.findById("wnd[0]/usr/ctxtLISTU").text = Scope_of_list[0]
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_CN_VBELN_%_APP_%-VALU_PUSH",value_list=Selection_parameter,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_SELPA_%_APP_%-VALU_PUSH",value_list=Document_type,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_BSART_%_APP_%-VALU_PUSH",value_list=Purchasing_group,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_EKGRP_%_APP_%-VALU_PUSH",value_list=Plant,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_WERKS_%_APP_%-VALU_PUSH",value_list=Item_category,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_PSTYP_%_APP_%-VALU_PUSH",value_list=Account_assignment_category,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_EINDT_%_APP_%-VALU_PUSH",value_list=Delivery_date,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_EBELN_%_APP_%-VALU_PUSH",value_list=Document_no ,ClipBoard_path=ouput_Path)

                        self.session.findById("wnd[0]/usr/ctxtP_GULDT").text = Validity_key_date[0]
                        self.session.findById("wnd[0]/usr/ctxtP_RWEIT").text = Range_to[0]
                                    
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_LIFNR_%_APP_%-VALU_PUSH",value_list=Vendor,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_RESWK_%_APP_%-VALU_PUSH",value_list=Delivering_plant,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH",value_list=Material,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_MATKL_%_APP_%-VALU_PUSH",value_list=Material_group,ClipBoard_path=ouput_Path)

                        
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_BEDAT_%_APP_%-VALU_PUSH",value_list=Document_date,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_EAN11_%_APP_%-VALU_PUSH",value_list=European_article_number,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_IDNLF_%_APP_%-VALU_PUSH",value_list=Vendor_material_number,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_LTSNR_%_APP_%-VALU_PUSH",value_list=Vendor_sub_range,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_AKTNR_%_APP_%-VALU_PUSH",value_list=Action,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_SAISO_%_APP_%-VALU_PUSH",value_list=Season,ClipBoard_path=ouput_Path)
                        self.set_transaction_parameter(id="wnd[0]/usr/btn%_S_SAISJ_%_APP_%-VALU_PUSH",value_list=Season_year,ClipBoard_path=ouput_Path)

                        self.session.findById("wnd[0]/usr/txtP_TXZ01").text = Short_text[0]
                        self.session.findById("wnd[0]/usr/txtP_NAME1").text = Vendor_name[0]
                        
      
                        
                        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()  # Run  <--------------------
                        time.sleep(1)
                        # ----------------------------------Validate transaction running ------------------------------------------------------
                        response=self.validate_run_response()
                        if not response: return response
                        time.sleep(3)


                        #--------------------------------------------------SET columns---------------------------------------------------
                        # if Columns[0] != "" :   
                        #         self.session.findById("wnd[0]/usr/cntlALVCONTAINER/shellcont/shell").pressToolbarContextButton ("&MB_VARIANT")
                        #         self.session.findById("wnd[0]/usr/cntlALVCONTAINER/shellcont/shell").selectContextMenuItem ("&COL0")

                        #         selected_cols_coun=0
                        #         for counter in range(0,105-len(Columns)):
                        #                 id="wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell"
                        #                 col=self.session.findById(id).text
                                        
                        #                 if col in Columns:
                        #                         self.session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press()
                        #                         selected_cols_coun=selected_cols_coun+1
                        #                         if selected_cols_coun == len(Columns): break
                        #                 else:
                        #                         #self.session.findById("wnd[1]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/tblSAPLCNFATC_ALLE_FELDER").verticalScrollbar.position = counter+1
                        #                         self.session.findById(id).currentCellRow = counter+1
                        #                         self.session.findById(id).firstVisibleRow = counter+1
                        #                         self.session.findById(id).selectedRows = str(counter+1)
                        #         self.session.findById("wnd[1]/tbar[0]/btn[0]").press()

                        
                        self.session.findById("wnd[0]/tbar[1]/btn[33]").press()
                        time.sleep(2)
                        self.session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cmbG51_USPEC_LBOX").SetFocus()
                        self.session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").contextMenu()
                        self.session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectContextMenuItem ("&FILTER")
                        time.sleep(2)        
                        self.session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = Layout
                        
                        self.session.findById("wnd[2]/tbar[0]/btn[0]").press()
                        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()                                                    

                        
                        #Set Full characteres len for all fields
                        # self.session.findById("wnd[0]/tbar[1]/btn[5]").press()
                        # self.session.findById("wnd[1]/usr/ssubRAHMEN:SAPLCNFA:0111/subAKT_FELDER:SAPLCNFA:0120/tblSAPLCNFATC_AKT_FELDER/txtAKT_FELDER-SCRTEXT[0,0]").setFocus()
                        # self.session.findById("wnd[1]/usr/ssubRAHMEN:SAPLCNFA:0111/subAKT_FELDER:SAPLCNFA:0120/tblSAPLCNFATC_AKT_FELDER/txtAKT_FELDER-SCRTEXT[0,0]").caretPosition = 11
                        # self.session.findById("wnd[1]/tbar[0]/btn[6]").press()
                        # self.session.findById("wnd[1]/usr/ssubRAHMEN:SAPLCNFA:0111/subANZAHL:SAPLCNFA:0150/subANZAHL:SAPLCNFA:0151/btnORIGINALLAENGE_HOLEN").press()
                        # self.session.findById("wnd[1]/tbar[0]/btn[7]").press()
                        # self.session.findById("wnd[1]/tbar[0]/btn[0]").press()

                        # Export File
                        #self.session.findById("wnd[0]/usr/cntlALVCONTAINER/shellcont/shell").pressToolbarContextButton ("&MB_EXPORT")
                        #self.session.findById("wnd[0]/usr/cntlALVCONTAINER/shellcont/shell").selectContextMenuItem ("&PC")
                        
                        self.session.findById("wnd[0]/tbar[1]/btn[45]").press()
                        #self.session.findById("wnd[0]/mbar/menu[0]/menu[11]/menu[6]").Select()
                        #self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select() <---un comment to get text with taps
                        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                        self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = ouput_Path
                        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = filename
                        self.session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
                        self.session.findById("wnd[1]/tbar[0]/btn[11]").press()

                        self.session.findById("wnd[0]/tbar[0]/btn[15]").press()
                        self.session.findById("wnd[0]/tbar[0]/btn[15]").press()
                        time.sleep(3)
                else:
                        print("Please Add values to extract")
                        
                return response
        
        #----------------------------------------------------------------------------------zpomon---------------------------------------------------------------------------------------

# if __name__=='__main__':
#         SAP_session = SapGui()
#         if not SAP_session.sap_Login(pUserName=user,pPassword=password): 
#                 exit

#         # print("-----Extract cji3")
#         # SAP_session.get_CJI3(PD=pd,ouput_Path=path,Posting_date=["03-04-2021","05-04-2022"])
#         # print("-----Extract cji5")
#         # SAP_session.get_CJI5(PD=pd,ouput_Path=path,Posting_date=["03-04-2021","05-04-2022"])
#         print("-----Extract cns41")
#         SAP_session.get_CNS41(PD=pd4,ouput_Path=path, Columns=["Project object"])
#         # print("-----Extract cns41")
#         # SAP_session.get_CNS41(PD=pd,ouput_Path=path, filename="CNS41_simple")
#         # print("-----Extract zsnap")
#         # SAP_session.get_ZSNAP(customer_number = CUSTOMER_NUMBER,customer_PO_number=["13943063"], Date_received=["03-04-2021","05-04-2022"],ouput_Path=path)
#         # print("-----Extract zzcustmon")
#         # SAP_session.get_Zzcustmon(Customer_PO=po_list,ouput_Path=path)
#         # print("-----Extract check")
#         # SAP_session.get_zcheckqtc(Project=['EUS07.VZ646','EUS03.NN296'],ouput_Path=path,layout="layout")
#         time.sleep(5)

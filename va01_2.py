import win32com.client as win32
import pythoncom
import win32com.client
import time

def va01_2(sesionsap, canal, sector, ped_ext, dispone, fecha_entrega, lista_id_productos, cantidades, convenio, lista_descripciones):

     pythoncom.CoInitialize()

     SapGuiAuto = win32com.client.GetObject('SAPGUI')
     if not type(SapGuiAuto) == win32com.client.CDispatch:
          return

     application = SapGuiAuto.GetScriptingEngine
     if not type(application) == win32com.client.CDispatch:
          SapGuiAuto = None
          return
     connection = application.Children(0)

     if not type(connection) == win32com.client.CDispatch:
          application = None
          SapGuiAuto = None
          return

     session = connection.Children(sesionsap)
     if not type(session) == win32com.client.CDispatch:
          connection = None
          application = None
          SapGuiAuto = None
          return

     try:
          #session.findById("wnd[0]").maximize()
          session.findById("wnd[0]/tbar[0]/okcd").text = "/NVA01"
          session.findById("wnd[0]").sendVKey(0)
          session.findById("wnd[0]/usr/ctxtVBAK-AUART").text = "ZTRA"
          session.findById("wnd[0]/usr/ctxtVBAK-VKORG").text = "SC10"
          session.findById("wnd[0]/usr/ctxtVBAK-VTWEG").text = canal
          session.findById("wnd[0]/usr/ctxtVBAK-SPART").text = sector
          session.findById("wnd[0]/usr/ctxtVBAK-SPART").setFocus()
          session.findById("wnd[0]/usr/ctxtVBAK-SPART").caretPosition = 2
          session.findById("wnd[0]").sendVKey(0)
          session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").text = ped_ext
          session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR").text = "20000123"
          session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUWEV-KUNNR").text = dispone
          session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").setFocus()
          session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").caretPosition = 8
          session.findById("wnd[0]").sendVKey(0)
          session.findById("wnd[0]").sendVKey(0)
          session.findById("wnd[0]").sendVKey(0)
                    
          session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/ctxtRV45A-KETDAT").text = fecha_entrega
          session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/ctxtRV45A-KETDAT").setFocus()
          session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/ctxtRV45A-KETDAT").caretPosition = 2
          session.findById("wnd[0]").sendVKey(0)

          session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02").select()

          try:
               for i in range(0, len(lista_id_productos)):
                    session.findById(f"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,{i}]").text = lista_id_productos[i]
                    session.findById(f"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[2,{i}]").text = cantidades[i]
                    
                    if lista_id_productos[i] == "99999":
                         session.findById(f"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-ARKTX[5,{i}]").text = lista_descripciones[i]
                    session.findById(f"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-WERKS[12,{i}]").text = "DSZA"
                    session.findById(f"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-LGORT[65,{i}]").text = "ALMA"
                    session.findById(f"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-LGORT[65,{i}]").setFocus()
                    session.findById(f"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-LGORT[65,{i}]").caretPosition = 4
                    session.findById("wnd[0]").sendVKey(0)
          except Exception as a:
               print("Linea 72 VA01-2", a)
               return
          
          session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press()
          session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\10").select()
          session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\12").select()
          session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13").select()
          session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/ctxtVBAK-ZZCONVENIO").text = convenio
          session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/ctxtVBAK-ZZCONVENIO").caretPosition = 2
          session.findById("wnd[0]").sendVKey(0)
          session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/ctxtVBAK-ZZTURNO").text = "MAN"
          session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/ctxtVBAK-ZZTURNO").setFocus()
          session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/ctxtVBAK-ZZTURNO").caretPosition = 3

          #Hace falta agregar una excepcion en este punto?
          session.findById("wnd[0]/tbar[0]/btn[11]").press()
          time.sleep(3)
          ped = session.findById("wnd[0]/sbar").text
          ped_final = ped[18:25]
          return ped_final

     except Exception as e:
          print("Linea 92 VA01-2", e)
          return -1
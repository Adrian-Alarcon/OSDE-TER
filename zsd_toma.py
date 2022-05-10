
from getpass import getuser
import win32com.client as win32
import pythoncom
import win32com.client
import time



def toma(sesionsap, ped_final, dispone, afiliado_sap, canal):
    
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
        session.findById("wnd[0]/tbar[0]/okcd").text = "/NZSD_TOMA"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/tbar[1]/btn[7]").press()
        session.findById("wnd[0]/usr/subSBS_PARSEL:ZDMSD_TOMA_PEDIDO:1100/ctxtS_VBELN-LOW").text = ped_final
        session.findById("wnd[0]/usr/subSBS_PARSEL:ZDMSD_TOMA_PEDIDO:1100/ctxtS_VBELN-LOW").setFocus()
        session.findById("wnd[0]/usr/subSBS_PARSEL:ZDMSD_TOMA_PEDIDO:1100/ctxtS_VBELN-LOW").caretPosition = 7
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/cntlCC_LISTAPED/shellcont/shell").currentCellColumn = "STAT_DISP_ICON"
        session.findById("wnd[0]/usr/cntlCC_LISTAPED/shellcont/shell").selectedRows = "0"
        session.findById("wnd[0]/usr/cntlCC_LISTAPED/shellcont/shell").pressToolbarButton("FN_MODPED")
        session.findById("wnd[0]/usr/tabsTABS/tabpTAB_PED/ssubTABS_SCA:ZDMSD_TOMA_PEDIDO:0101/cmbZSD_TOMA_CABEC-LIFSK").key = "00"
        session.findById("wnd[0]/usr/tabsTABS/tabpTAB_ENT").select()
        session.findById("wnd[0]/usr/tabsTABS/tabpTAB_ENT/ssubTABS_SCA:ZDMSD_TOMA_PEDIDO:0102/ctxtGS_ENTREGA-AFIL_NRO").text = afiliado_sap
        
        if canal != "10":
            session.findById("wnd[0]/usr/tabsTABS/tabpTAB_ENT/ssubTABS_SCA:ZDMSD_TOMA_PEDIDO:0102/ctxtGS_ENTREGA-DISPONE_ID").text = dispone
        
        session.findById("wnd[0]/usr/tabsTABS/tabpTAB_ENT/ssubTABS_SCA:ZDMSD_TOMA_PEDIDO:0102/ctxtGS_ENTREGA-DISPONE_ID").setFocus()
        session.findById("wnd[0]/usr/tabsTABS/tabpTAB_ENT/ssubTABS_SCA:ZDMSD_TOMA_PEDIDO:0102/ctxtGS_ENTREGA-DISPONE_ID").caretPosition = 8
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/tabsTABS/tabpTAB_ENT/ssubTABS_SCA:ZDMSD_TOMA_PEDIDO:0102/btnBTN_CALC_FECHA").press()
        session.findById("wnd[0]/usr/tabsTABS/tabpTAB_ENT/ssubTABS_SCA:ZDMSD_TOMA_PEDIDO:0102/btnBTN_CALC_FECHA").press()
        session.findById("wnd[0]/usr/tabsTABS/tabpTAB_ENT/ssubTABS_SCA:ZDMSD_TOMA_PEDIDO:0102/btnBTN_CALC_FECHA").press()
        session.findById("wnd[0]/usr/tabsTABS/tabpTAB_PED").select()
        session.findById("wnd[0]/usr/tabsTABS/tabpTAB_PED/ssubTABS_SCA:ZDMSD_TOMA_PEDIDO:0101/subSUBS_TRAB:ZDMSD_TOMA_PEDIDO:0111/btnBTN_SIMULAR").press()
        
        try:
            session.findById("wnd[0]/tbar[0]/btn[11]").press()
            session.findById("wnd[1]/usr/btnBUTTON_1").press()
        except:    
            session.findById("wnd[1]/usr/btnBUTTON_1").press()
        
        try:
            session.findById("wnd[1]/usr/btnBUTTON_1").press()
        except:
            try:
                ##sincartel11:
                session.findById("wnd[1]/usr/btnBUTTON_1").press()
            except:
                ##sinvalidacion:
                session.findById("wnd[1]/tbar[0]/btn[0]").press()
        try:
            session.findById("wnd[0]/tbar[1]/btn[7]").press()
            session.findById("wnd[1]").sendVKey(0)
            session.findById("wnd[1]").sendVKey(0)
        except:
            session.findById("wnd[1]").sendVKey(0)
            session.findById("wnd[1]").sendVKey(0)
            return ped_final
        return ped_final
    
    except:
        time.sleep(3)
        return f"{ped_final}"

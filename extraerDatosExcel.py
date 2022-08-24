import time
import rutas
from openpyxl import load_workbook
import os
from va01_2 import va01_2
from zsd_toma import toma
import shutil
import pythoncom


def cargar_listas_datos(hoja, cant_filas):
    """
    ESTA FUNCION SE ENCARGA DE RECIBIR LA HOJA DEL EXCEL Y LA CANTIDAD DE FILAS
    PARA PODER CARGAR TODOS LOS DATOS RELEVANTES A LAS LISTAS
    """
    l_mat_cliente_e = []
    l_cant_e = []
    l_ped_ext_e = []
    l_observ_e = []
    l_dispones_e = []
    l_id_afil_e = []
    l_id_mat_sap_e = []
    l_convenios_e = []
    fila_ped_cargado = []

    i = 2
    while i <= cant_filas:
        # REVISAR SI LA FILA SE PUEDE CARGAR (revision = NO)
        revision = hoja[f"H{i}"].value
        if revision == "NO":
            # AGREGAR DATOS RELEVANTES A LAS LISTAS
            l_mat_cliente_e.append(hoja[f"D{i}"].value)
            l_cant_e.append(hoja[f"E{i}"].value)
            l_ped_ext_e.append(hoja[f"F{i}"].value)
            l_observ_e.append(hoja[f"G{i}"].value)
            l_dispones_e.append(hoja[f"N{i}"].value)
            l_id_afil_e.append(hoja[f"O{i}"].value)
            l_id_mat_sap_e.append(hoja[f"P{i}"].value)
            l_convenios_e.append(hoja[f"Q{i}"].value)
            fila_ped_cargado.append(str(i))
        else:
            continue
        i += 1
    return l_mat_cliente_e, l_cant_e, l_ped_ext_e, l_observ_e, l_dispones_e, l_id_afil_e, l_id_mat_sap_e, l_convenios_e, fila_ped_cargado


def cargar_pedidos(tupla_lista_datos, hoja):
    l_mat_cliente_e,\
    l_cant_e, l_ped_ext_e, \
    l_observ_e, \
    l_dispones_e, \
    l_id_afil_e, \
    l_id_mat_sap_e, \
    l_convenios_e, \
    fila_ped_cargado = tupla_lista_datos

    mat_cl_cargar = []
    cantidades_cargar = []
    filas_completar = []
    mat_sap_cargar = []
    afiliado_anterior = None

    for d in range(len(afi_osde)):
        afiliado_actual = afi_osde[d]
        if afiliado_actual == afiliado_anterior or afiliado_anterior == None:
            mat_cl_cargar.append(id_mat_cliente[d])
            mat_sap_cargar.append(id_mat_sap[d])
            cantidades_cargar.append(cantidades[d])
            filas_completar.append(fila_ped_cargado[d])
        elif afiliado_anterior != afiliado_actual:
            try:
                print(f"Se factura: {afiliado_anterior}")
                print(f"{afiliado_anterior}, {mat_cl_cargar}, {cantidades_cargar}, {ped_externo[d-1]}," +
                      f"{observ_internas[d-1]}, {dispones[d-1]}, {id_afil_sap[d-1]}, {convenio[d-1]}," +
                      f"{fecha[d-1]}, {filas_completar}")
                ped_va = va01_2(0, ped_externo[d-1], dispones[d-1], fecha[d], mat_sap_cargar, cantidades_cargar, convenio[d-1], mat_cl_cargar)
                ped_toma = toma(0, ped_va, dispones[d-1], id_afil_sap[d-1], "02", observ_internas[d-1])

                for fila in filas_completar:
                    hoja[f"AA{int(fila)}"].value = ped_toma

                print("-------------------------------------------------------------------------------------")
                mat_cl_cargar.clear()
                mat_sap_cargar.clear()
                cantidades_cargar.clear()
                filas_completar.clear()
                mat_cl_cargar.append(id_mat_cliente[d])
                mat_sap_cargar.append(id_mat_sap[d])
                cantidades_cargar.append(cantidades[d])
                filas_completar.append(fila_ped_cargado[d])
            except Exception as e:
                print(f"Error en VA01 {e}")
        elif d == len(afi_osde):
            print(f"Se factura: {afiliado_anterior}")
            print(f"{afiliado_anterior}, {mat_cl_cargar}, {cantidades_cargar}, {ped_externo[d]}," +
                  f"{observ_internas[d]}, {dispones[d]}, {id_afil_sap[d]}, {convenio[d]}," +
                  f"{fecha[d]}, {filas_completar}")
            ped_va = va01_2(0, ped_externo[d], dispones[d], fecha[d], mat_sap_cargar, cantidades_cargar, convenio[d], mat_cl_cargar)
            ped_toma = toma(0, ped_va, dispones[d], id_afil_sap[d], "02", observ_internas[d])

            for fila in filas_completar:
                hoja[f"AA{int(fila)}"].value = ped_toma

            print("-------------------------------------------------------------------------------------")
            break
        afiliado_anterior = afi_osde[d]


# --- PROGRAMA PRINCIPAL --- #
def carga_sap():
    pythoncom.CoInitialize()
    shutil.copy(rutas.archivo_excel, rutas.archivo_excel_trabajo)
    time.sleep(2)
    existe_ruta = os.path.exists(rutas.archivo_excel_trabajo)
    if existe_ruta:
        try:
            excel = load_workbook(rutas.archivo_excel_trabajo, data_only=True)
            hoja = excel["inicio"]
            cant_filas = int(hoja["A2"].value)

            # FUNCION QUE CARGA LA INFORMACION DEL EXCEL Y DEVUELVE LISTAS CARGADAS DE DATOS
            tupla_lista_de_datos = cargar_listas_datos(hoja, cant_filas)
            for listas in tupla_lista_de_datos:
                print(listas)

            # cargar_pedidos(tupla_lista_de_datos, hoja)

        except Exception as e:
            print(e)
        finally:
            excel.save(rutas.archivo_excel_trabajo)
    else:
        print("El Excel no existe! Revisar...")


carga_sap()
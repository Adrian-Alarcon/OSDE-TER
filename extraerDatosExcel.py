import rutas
from openpyxl import load_workbook
import os.path
from va01_2 import va01_2
from zsd_toma import toma


def cargar_listas_datos(hoja, cant_filas):
    """
    ESTA FUNCION SE ENCARGA DE RECIBIR LA HOJA DEL EXCEL Y LA CANTIDAD DE FILAS
    PARA PODER CARGAR TODOS LOS DATOS RELEVANTES A LAS LISTAS
    """
    afi_osde = []
    id_mat_cliente = []
    cantidades = []
    id_mat_sap = []
    ped_externo = []
    observ_internas = []
    dispones = []
    id_afil_sap = []
    convenio = []
    fecha = []
    fila_ped_cargado = []

    i = 2
    while i <= cant_filas:
        # REVISAR SI LA FILA SE PUEDE CARGAR (revision = NO)
        revision = hoja[f"H{i}"].value
        if revision == "NO":
            # AGREGAR DATOS RELEVANTES A LAS LISTAS
            afi_osde.append(hoja[f"B{i}"].value)
            id_mat_cliente.append(hoja[f"D{i}"].value)
            cantidades.append(hoja[f"E{i}"].value)
            ped_externo.append(hoja[f"F{i}"].value)
            observ_internas.append(hoja[f"G{i}"].value)
            dispones.append(hoja[f"N{i}"].value)
            id_afil_sap.append(hoja[f"O{i}"].value)
            convenio.append(hoja[f"Q{i}"].value)
            fecha.append(hoja[f"V{i}"].value)
            fila_ped_cargado.append(str(i))
        else:
            continue
        i += 1
    return afi_osde, id_mat_cliente, cantidades, ped_externo, observ_internas, dispones, id_afil_sap, convenio, fecha, fila_ped_cargado


def cargar_pedidos(tupla_lista_datos):
    afi_osde, id_mat_cliente, cantidades, ped_externo, observ_internas, dispones, id_afil_sap, convenio, fecha, fila_ped_cargado = tupla_lista_datos

    mat_cl_cargar = []
    cantidades_cargar = []
    filas_completar = []
    afiliado_anterior = None

    for d in range(len(afi_osde)):
        afiliado_actual = afi_osde[d]
        if afiliado_actual == afiliado_anterior or afiliado_anterior == None:
            mat_cl_cargar.append(id_mat_cliente[d])
            cantidades_cargar.append(cantidades[d])
            filas_completar.append(fila_ped_cargado[d])
        elif afiliado_anterior != afiliado_actual:
            try:
                print(f"Se factura: {afiliado_anterior}")
                print(f"{afiliado_anterior}, {mat_cl_cargar}, {cantidades_cargar}, {ped_externo[d-1]}," +
                      f"{observ_internas[d-1]}, {dispones[d-1]}, {id_afil_sap[d-1]}, {convenio[d-1]}," +
                      f"{fecha[d-1]}, {filas_completar}")
                print("-------------------------------------------------------------------------------------")
                mat_cl_cargar.clear()
                cantidades_cargar.clear()
                filas_completar.clear()
                mat_cl_cargar.append(id_mat_cliente[d])
                cantidades_cargar.append(cantidades[d])
                filas_completar.append(fila_ped_cargado[d])
            except Exception as e:
                print(f"Error en VA01 {e}")
        elif d == len(afi_osde):
            print(f"Se factura: {afiliado_anterior}")
            print(f"{afiliado_anterior}, {mat_cl_cargar}, {cantidades_cargar}, {ped_externo[d]}," +
                  f"{observ_internas[d]}, {dispones[d]}, {id_afil_sap[d]}, {convenio[d]}," +
                  f"{fecha[d]}, {filas_completar}")
            print("-------------------------------------------------------------------------------------")
            break
        afiliado_anterior = afi_osde[d]


# --- PROGRAMA PRINCIPAL --- #
existe_ruta = os.path.exists(rutas.archivo_excel_trabajo)

if existe_ruta:
    try:
        excel = load_workbook(rutas.archivo_excel_trabajo)
        hoja = excel["inicio"]
        cant_filas = int(hoja["A2"].value)

        # FUNCION QUE CARGA LA INFORMACION DEL EXCEL Y DEVUELVE LISTAS CARGADAS DE DATOS
        tupla_lista_de_datos = cargar_listas_datos(hoja, cant_filas)
        print(tupla_lista_de_datos)
        # CARGAR LOS PEDIDOS CON LOS DATOS PREVIAMENTE CARGADOS EN LAS LISTAS
        cargar_pedidos(tupla_lista_de_datos)

    except Exception as e:
        print(e)
    finally:
        excel.save(rutas.archivo_excel_trabajo)

else:
    print("El Excel no existe! Revisar...")

#print(afi_osde, id_mat_cliente, cantidades, ped_externo, observ_internas, dispones, id_afil_sap, convenio, fecha, fila_ped_cargado, sep="\n")

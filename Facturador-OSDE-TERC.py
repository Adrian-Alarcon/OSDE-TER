import rutas
from va01_2 import va01_2
from zsd_toma import toma
from tkinter import *
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from getpass import getuser
import shutil
import pythoncom
from openpyxl import load_workbook
from time import sleep


def interfaz():

     user = getuser()
    
     root = Tk()
     
     root.title("CARGA-OSDE-TERC")
     root.resizable(0,0)
     root.geometry('200x200+300+20'.format(300, 200))

     miFrame=Frame(root,width=500)
     miFrame.pack()
     miFrame2=Frame(root)
     miFrame2.pack()

     def facturador():

          pythoncom.CoInitialize()
          # Crear una copia del excel padre.
          shutil.copy(rutas.archivo_excel, rutas.archivo_excel_trabajo)

          #---------VARIABLES--------#
          afiliado_anterior = None
          afiliado_actual = None

          l_obsersevaciones_int = []
          l_id_productos = []
          l_cantidades_va01 = []
          l_afiliados_sap = []
          l_dispones = []
          l_canales = []
          l_sectores = []
          l_ped_ext = []
          l_convenios = []
          l_mat_cliente = []
          l_tipo_dispone = [] # Para identificar los centros asistenciales
          l_ce_asist = []

          l_mat_cliente_fact = []
          l_observ_toma = []
          l_mat_facturar = []
          l_cant_facturar = []
          fecha_entrega = []
          filas_completar = []
          tipo_dispone = [] # Para identificar los centros asistenciales
          centro_asistencial = []
          l_filas = []
#---------------------------#

          excel_trabajo = load_workbook(rutas.archivo_excel_trabajo, data_only=True)
          try:
               h_t = excel_trabajo["inicio"]

               # === CARGAR LISTAS === #
               cont = 2
               max_filas = int(h_t[f"A2"].value)
               for i in range(2, max_filas + 1):
                    revision = h_t[f"H{i}"].value
                    if revision == "SI":
                         print(f"AFILIADO {afiliado_actual} para REVISION")
                         continue
                    else:
                         l_afiliados_sap.append(h_t[f"O{i}"].value)
                         l_id_productos.append(h_t[f"P{i}"].value)
                         l_obsersevaciones_int.append(h_t[f"G{i}"].value)
                         l_mat_cliente.append(h_t[f"D{i}"].value)
                         l_cantidades_va01.append(h_t[f"E{i}"].value)
                         l_canales.append(h_t[f"Y{i}"].value)
                         l_sectores.append(h_t[f"Z{i}"].value)
                         l_ped_ext.append(h_t[f"F{i}"].value)
                         l_dispones.append(h_t[f"N{i}"].value)
                         fecha_entrega.append(h_t[f"V{i}"].value)
                         l_convenios.append(h_t[f"Q{i}"].value)
                         l_tipo_dispone.append(h_t[f"S{i}"].value)
                         l_ce_asist.append(h_t[f"T{i}"].value)
                         filas_completar.append(str(i))
               print(f"Filas a completar:", filas_completar)


               # ==== CARGA DE PEDIDOS ==== #
               for i in range(len(l_afiliados_sap)):
                    afiliado_actual = l_afiliados_sap[i]

                    if afiliado_actual == afiliado_anterior or afiliado_anterior == None:
                         l_observ_toma.append(l_obsersevaciones_int[i])
                         l_mat_facturar.append(l_id_productos[i])
                         l_mat_cliente_fact.append(l_mat_cliente[i])
                         l_cant_facturar.append(l_cantidades_va01[i])
                         tipo_dispone.append(l_tipo_dispone[i])
                         centro_asistencial.append(l_ce_asist[i])
                         l_filas.append(filas_completar[i])

                    elif afiliado_actual != afiliado_anterior:
                         print(f"SE CARGARA PED: {afiliado_anterior} - {i} - {tipo_dispone} - {centro_asistencial}")
                         print("\tYENDO A VA01:", l_ped_ext[i-1], l_dispones[i-1], fecha_entrega[i-1], l_mat_facturar, l_cant_facturar, l_convenios[i-1], l_mat_cliente_fact, tipo_dispone[0], centro_asistencial[0])

                         pedidova01 = va01_2(0, l_ped_ext[i-1], l_dispones[i-1], fecha_entrega[i-1], l_mat_facturar, l_cant_facturar, l_convenios[i-1], l_mat_cliente_fact, tipo_dispone[0], centro_asistencial[0], l_canales[i-1], l_sectores[i-1])

                         sleep(1)
                         if pedidova01 != -1:
                              print("\tYENDO A TOMA:", f"{pedidova01}", l_dispones[i-1], afiliado_anterior, l_observ_toma)
                              _toma = toma(0, pedidova01, afiliado_anterior, l_dispones[i-1], l_observ_toma)
                              #Completar Excel con pedido generado en TOMA
                              for fila in l_filas:
                                   h_t[f"AA{fila}"] = _toma
                         else:
                              print(f"NO SE PUDO CARGAR PEDIDO EN VA01 PARA AF: {afiliado_anterior}")
                              # Completar Excel con pedido generado en VA01
                              for fila in l_filas:
                                   h_t[f"AA{fila}"] = pedidova01 + "cargado solamente en VA01"

                         l_observ_toma.clear()
                         l_mat_facturar.clear()
                         l_mat_cliente_fact.clear()
                         l_cant_facturar.clear()
                         tipo_dispone.clear()
                         centro_asistencial.clear()
                         l_filas.clear()

                         l_observ_toma.append(l_obsersevaciones_int[i])
                         l_mat_facturar.append(l_id_productos[i])
                         l_mat_cliente_fact.append(l_mat_cliente[i])
                         l_cant_facturar.append(l_cantidades_va01[i])
                         tipo_dispone.append(l_tipo_dispone[i])
                         centro_asistencial.append(l_ce_asist[i])
                         l_filas.append(filas_completar[i])

                    if i == len(l_afiliados_sap) - 1:
                         print(f"{i} - ULTIMO AFILIADO: {afiliado_actual}")
                         print("\tULTIMA VUELTA VA01:",0, l_ped_ext[i], l_dispones[i], fecha_entrega[i], l_mat_facturar, l_cant_facturar, l_convenios[i], l_mat_cliente_fact, tipo_dispone[0], centro_asistencial[0])

                         pedidova01 = va01_2(0, l_ped_ext[i], l_dispones[i], fecha_entrega[i], l_mat_facturar, l_cant_facturar, l_convenios[i], l_mat_cliente_fact, tipo_dispone[0], centro_asistencial[0], l_canales[i], l_sectores[i])

                         sleep(1)
                         if pedidova01 != -1:
                              print("\tULTIMA VUELTA TOMA:",0, "pedidova01", l_dispones[i], afiliado_actual, l_observ_toma)
                              _toma = toma(0, pedidova01, afiliado_actual, l_dispones[i], l_observ_toma)
                              #Completar Excel con pedido generado
                              for fila in l_filas:
                                   h_t[f"AA{fila}"] = _toma
                         else:
                              print(f"NO SE PUDO CARGAR PEDIDO EN VA01 PARA AF: {afiliado_anterior}")
                              for fila in l_filas:
                                   h_t[f"AA{fila}"] = pedidova01 + "cargado solamente en VA01"

                    print()
                    afiliado_anterior = afiliado_actual

          except Exception as e:
               print(f"Excepcion el Excel de Trabajo {e}")
          finally:
               excel_trabajo.save(rutas.archivo_excel_trabajo)
               excel_trabajo.close()

     botonCrear = Button(miFrame, text="Ejecutar", command=facturador)
     botonCrear.grid(row = 13, column = 2, sticky = "e", padx = 15, pady = 40)

     root.mainloop()

interfaz()
import os
from getpass import getuser
from datetime import datetime

now = datetime.now()

#USUARIO LOCAL
usuario = getuser()

# --> obtenemos la ruta del directorio actual.
directorio_raiz = f"C:/Users/{usuario}/Documents/Clientes Carga Automatica/OSDE_TER"

# --> Armamos la ruta del directorio de archivos.
directorio_archivos = directorio_raiz + "/" + "archivos"

# --> Armamos la ruta del Excel.
archivo_excel_trabajo = directorio_archivos + "/" + "BaseOsdeTer.xlsx"
archivo_excel = directorio_archivos + "/" + "ArchivoTrabajoOsdeTer -- " + now.strftime('%m-%d-%Y') + ".xlsx"




import pandas as pd
from tkinter import messagebox

def crear_nuevo_excel(excel_cias_parcial, excel_global , excel_resultante):
    """_summary_

    Args:
        excel_cias_parcial (_type_): Excel con la lista de destinatarios parcial sin sus correos.
        excel_global: Excel con la totalidad de las cias y sus correos.
        excel_totalidad_cias (_type_): Es el nombre con el que se guardara el excel final.

    Raises:
        ValueError: _description_

    Returns:
        _type_: Un excel mergeado con las cias 
    """
    try:
        # ------------- LEER EL EXCEL DE CIAS PARCIAL -------------

        # Determinar la extensión del archivo
        file_extension = excel_cias_parcial.split('.')[-1]
        
        # Cargar los archivos Excel según la extensión
        if file_extension == 'xls':
            df1 = pd.read_excel(excel_cias_parcial, engine='xlrd')
        elif file_extension == 'xlsx':
            df1 = pd.read_excel(excel_cias_parcial, engine='openpyxl')
        else:
            raise ValueError("Formato de archivo no soportado. Use 'xls' o 'xlsx'.")
        
        # ------------- LEER EL EXCEL DE CIAS TOTALIDAD -------------
        # excel_global es el excel base con todos los destinatarios y sus correos
        
        df2 = pd.read_excel(f'./excels/{excel_global}.xlsx', engine='openpyxl')

        # Seleccionar solo la columna "CORREOS" junto con "CIA" para el merge
        df2 = df2[['CIA', 'CORREOS']]

        # # ------------- FUSIONAR AMBOS EXCELS -------------
        
        # Fusionar los dataframes en la columna "Cia ID"/"CIA"
        merged_df = pd.merge(df1, df2, left_on='Cia ID', right_on='CIA', how='left')

        # Eliminar la columna "CIA" del dataframe resultante
        merged_df = merged_df.drop(columns=['CIA'])

        # Guarda el dataframe fusionado en un nuevo archivo Excel
        merged_df.to_excel(excel_resultante, index=False)

        # ---------------------------------------------------

        print(f"Excel {excel_resultante} generado con exito")
        messagebox.showinfo("Exito", f"Excel {excel_resultante} generado con exito")

        return True
    
    except Exception as e:
        print(f"Error en la union de los excels: {e}")
        messagebox.showerror("Error", f"Error en la union de los excels '{excel_cias_parcial}' y '{excel_global}'")
        return False





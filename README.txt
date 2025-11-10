INSTRUCCIONES PARA EJECUTAR LA APLICACIÓN WEB CON STREAMLIT CLOUD

1. Requisitos:
   - Navegador web Edge o Chrome.

2. Ejecución:
   - Abrir página web:    https://informes-rectauto.streamlit.app/

3. Uso:
   - Cargar los archivos requeridos:
         - RECTAUTO, tal y como se descarga de Qlik Sense en formato Excel.
         - NOTIFICA, tal y como se descarga de Qlik Sense en formato Excel.
         - USUARIOS, hoja de cálculo que requiere mantenimiento cuando haya
            que modificar datos de los usuarios (altas, bajas, IT de larga
            duración).
         - TRIAJE, hoja de cálculo  mantenida en local con las asignaciones
            semanales (solo el último archivo, ya que todos los expedientes
            actualmente abiertos, están en esa hoja.
         - DOCUMENTOS, hoja de cálculo cuyo mantenimiento se realiza en la
            misma app-web, ya que grabamos los datos actualizados de esos
            expedientes, y descargamos el fichero que se utilizará la siguien-
            vez que se utilice la app-web.
         - Se obtienen los informes individuales, los de Expedientes Priori-
            tarios de los equipos y el resumen de KPI semanal, descargándose 
            en un archivo comprimido .zip.
         - En la app-web NO SE PUEDE REALIZAR EL ENVÍO DE CORREO DIRECTAMENTE.
            En la app de escritorio sí se puede, teniendo Outlook instalado.

Control de cambio Agente Wells Fargo para agregar clientes al archivo de cartera que no tienen documentos abiertos.

**Objetivo:** Cambiar fuente de datos de los clientes, ahora el agente tendrá que hacer login en SAP ERP GUI, entrar a una transacción, extraer un reporte y guardarlo para continuar con el proceso.

**Pasos:**

1. Abrir SAP ERP GUI y hacer login con la conexión Componentes ERP/1. Grupo Nutresa\_ERP\_PRD, las credenciales que se proveerán, el password cambia cada 30 días, debemos establecer una forma sencilla para actualizar la clave para el agente.  
2. Entrar a la transacción ZSD\_POS\_1052, obtener la variante de parámetros CUSA-WF, ejecutar la transacción.  
3. Exportar a excel el resultado en el path customizable del agente, el nombre del archivo debe de ser siempre Customers\_SAP.XLSX, sobrescribir siempre.  
4. Salir de transacción y hacer logout en SAP.  
5. Se adjuntará un video del proceso antes descrito.  
6. El archivo cuenta con 3 columnas, los clientes a buscar si existen en el archivo FlatFile.csv deben ser solamente los que tienen el valor C030 en la tercer columna con el encabezado Terms of payment, seguir misma lógica del agente, si no se encuentra entonces insertar con las especificaciones antes dadas. A continuación nuevamente el mapeo de los campos para la inserción ahora con la nueva fuente:  
   1. Payer Code: Columna Customer de el archivo Customers\_SAP.XLSX, importante remover los primeros dos ceros a la izquierda del codigo del cliente, debe de quedar de 8 números (si son letras pues solo las letras)  
   2. Auth: Columna Customer de el archivo Customers\_SAP.XLSX, aplicar misma lógica del paso a.  
   3. Customer: Columna Name 1 de el archivo Customers\_SAP.XLSX  
   4. Due Date: la fecha de hoy con formato "%m/%d/%Y"  
   5. Amount Due: Valor cero  
   6. Biller Invoice No.: Dato fijo 99-99999999  
   7. P.O.: POAdvance  
   8. Customer Name: Columna Name 1 de el archivo Customers\_SAP.XLSX  
   9. Bank Account: Dato fijo 4942472523

El resto del procedimiento continúa igual como se definió en la especificación pasada.


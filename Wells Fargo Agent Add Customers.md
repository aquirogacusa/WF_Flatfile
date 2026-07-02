Especificación Agente Wells Fargo para agregar clientes al archivo de cartera que no tienen documentos abiertos.

**Objetivo:** Actualmente enviamos un archivo de cuentas por cobrar a nuestro banco, con todos los documentos de cuentas por cobrar activas. Tenemos la necesidad de que en dicho archivo aparezcan el 100% de nuestros clientes activos, tengan o no tengan documentos de cuentas por cobrar activas, por lo que se requiere revisar que clientes no aparecen en el archivo de cuentas por cobrar y agregarlos con un documento “dummy” con valor cero.

**Pasos:**

1. Descomprimir el archivo de cuentas por cobrar que tiene el nombre Wells Fargo Bill File.zip , dentro de el esta el archivo FlatFile.csv que es nuestro archivo de cuentas por cobrar.  
2. El archivo Customers\_WFEBill.xlsx tiene el 100% de los clientes activos, debemos de revisar que clientes que estan en el archivo Customers\_WFEBill.xlsx no se encuentran presentes en algún registro del archivo FlatFile.csv, la llave de busqueda es para el caso del archivo FlatFile.csv, la segunda columna con el encabezado Auth, para el caso del archivo Customers\_WFEBill.xlsx es la primera columna con encabezado Código Cliente.  
3. Una vez identificados los clientes que no se encuentran presentes en el archivo FlatFile.csv, se procede a insertarlos al final en el archivo FlatFile.csv, a continuación el detalle de lo que debe de ir en cada columna:  
   - Payer Code: Columna Código Cliente de el archivo Customers\_WFEBill.xlsx  
   - Auth: Columna Código Cliente de el archivo Customers\_WFEBill.xlsx  
   - Customer: Columna Nombre Cliente de el archivo Customers\_WFEBill.xlsx  
   - Due Date: la fecha de hoy con formato "%m/%d/%Y"  
   - Amount Due: Valor cero  
   - Biller Invoice No.: Dato fijo 99-99999999  
   - P.O.: POAdvance  
   - Customer Name: Columna Nombre Cliente de el archivo Customers\_WFEBill.xlsx  
   - Bank Account: Dato fijo 4942472523  
4. Una vez terminada la inserción de los datos, hay que eliminar el archivo Wells Fargo Bill File.zip y volver a crear un archivo zip del mismo nombre con el archivo FlatFile.csv dentro.Tambien eliminar el archivo FlatFile.csv que extrajimos en el paso 1\.

Debemos preparar el agente para que se pueda customizar el path donde estarán todos los archivos mencionados anteriormente.


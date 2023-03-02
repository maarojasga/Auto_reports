# DOCUMENTACIÓN CÓDIGO AUTOMATIZACIÓN REPORTES POS

### Directorio:

- **Carperta principal**
    - Carpeta del proyecto
    - env (Carpeta del entorno virtual)

### Pasos: (Comandos)

1. **Creación y activación del entorno virtual.**\
    python3 -m venv -env #Crea un entorno virtual llamado env\
    env\Scripts\activate #Activa el entorno virtual 

2. **Instalación de los requerimientos (Acceder a la carpeta del proyecto)**\
    pip install -r requirements.txt #Se instalan las librerías necesarias

3. **Creación de .env (Variables de entorno)**\
    Crear un archivo, a la misma altura de la carpeta del proyecto, llamado .env con las variables de entorno del proyecto.

4. **Ejecutar el código.**\
    python ReportesPOS\InsertData.py #Insertar datos en la tabla.\
    python ReportesPOS\SendData.py #Consultar vistas según la marca y enviar correos con los reportes en formato excel.


# NOTAS !IMPORTANT

1. **Tabla Email_Vendor:**\
    - La columna 'Subject' debe contener el asunto personalizado.
    - La columna 'Email' debe contener los correos separados por comas sin dejar espacios.
    - La columna 'Text'  debe contener el mensaje personalizado.


## Ejecutar tarear al completar con éxito otra tarea programada.

## 1. Begin the task_ On an event
## 2. Custom - New Event Filter
## 3. XMIL - Edit query manually
## 4. Write query --> 
<QueryList>
   <Query Id="0" Path="Microsoft-Windows-TaskScheduler/Operational">
      <Select Path="Microsoft-Windows-TaskScheduler/Operational">*[EventData[@Name='ActionSuccess'][Data [@Name='TaskName']='\FirstTask']] and *[EventData[@Name='ActionSuccess'][Data [@Name='ResultCode']='0']]</Select>
   </Query>
</QueryList>


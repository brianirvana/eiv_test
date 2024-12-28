# Documentación del Proyecto

## Descripción General
Este proyecto es una aplicación de escritorio desarrollada en **Visual Basic 6.0** que incluye funcionalidades clave como un sistema de login de usuarios, un ABM (Alta, Baja y Modificación) de personas y manejo integral de bases de datos. La base de datos subyacente es **MySQL**, y se implementó un conjunto de herramientas para profesionalizar el desarrollo.

## Funcionalidades Principales

### 1. Panel de Login de Usuarios
- **Botones**:
  - **Logear**: Permite a los usuarios ingresar al sistema después de autenticar sus credenciales.
  - **Crear Usuario**: Redirige al formulario de creación de nuevos usuarios.
- **Validaciones**:
  - Nombre de usuario: Debe tener al menos 3 letras.

### 2. Panel de Creación de Usuarios
- Validaciones implementadas:
  - **Nombre** y **Apellido**.
  - **Tipo de Documento**: Valida si se elije DNI u otro tipo definido.
  - **DNI**: Verifica que sea único.
  - **Email**: Valida el formato correcto de la dirección.
  - Otros campos relacionados con la información de la persona.

### 3. ABM de Personas (en curso)
Permite realizar operaciones de:
- **Alta**: Crear nuevas personas en la base de datos.
- **Baja**: Eliminar registros existentes.
- **Modificación**: Actualizar información de personas.

### 4. Funcionalidades de Base de Datos
- **Detección Automática**:
  - Si la base de datos no existe, se crea automáticamente.
  - Una vez creada, también se crean todas las tablas necesarias.
  - Si las tablas se crean, se inicializan con datos semilla.
- **Conexión Persistente**:
  - Si la base de datos existe, el sistema se conecta y mantiene la conexión de manera global a través del objeto `CN` (de tipo `ADODB.Connection`), lo que permite interactuar con la base de datos sin necesidad de reconectarse antes de cada consulta.

### 5. Creación Automática de Tablas
- La función `CreateTables` automatiza la creación de tablas siguiendo las mejores prácticas y asegurando la integridad referencial mediante claves foráneas.
- Las tablas creadas incluyen:
  - **tipos_documentos**: Define los tipos de documentos.
  - **provincias** y **localidades**: Manejo geográfico.
  - **personas**: Datos personales con referencia a documentos y localidades.
  - **usuarios**: Almacena credenciales y referencia a personas.

## Herramientas Utilizadas
- **ChatGPT**: Para el desarrollo de soluciones, refinamiento de código y generación de ideas.
- **MZ-Tools**: Utilizado para la generación de manejadores de errores y optimización de desarrollo en VB6.
- **Smart Indent**: Herramienta para indentar el código de forma eficiente, mejorando la legibilidad.

## Estructura de la Base de Datos

1. **tipos_documentos**
   - `id_tipodocumento`: Identificador único.
   - `nombre`: Nombre del tipo de documento.
   - `abreviatura`: Abreviatura del documento.
   - `validar_como_cuit`: Indicador de validación como CUIT.

2. **provincias**
   - `id_provincia`: Identificador único.
   - `nombre`: Nombre de la provincia.
   - `region`: Código de la región.

3. **localidades**
   - `id_localidad`: Identificador único.
   - `nombre`: Nombre de la localidad.
   - `id_provincia`: Referencia a `provincias`.
   - `codigo_postal`: Código postal.

4. **personas**
   - `id_persona`: Identificador único.
   - `id_tipodocumento`: Referencia a `tipos_documentos`.
   - `num_documento`: Número del documento.
   - `nombre_apellido`: Nombre y apellido.
   - `fecha_nacimiento`: Fecha de nacimiento.
   - `genero`: Género.
   - `correo_electronico`: Email.
   - `id_localidad`: Referencia a `localidades`.

5. **usuarios**
   - `id_persona`: Referencia a `personas`.
   - `nombre_usuario`: Nombre de usuario único.
   - `hashed_pwd`: Contraseña encriptada.

## Ejecución del Proyecto
1. **Configuración Inicial**:
   - Configurar la cadena de conexión en el archivo del proyecto.
2. **Ejecución Automática**:
   - Al iniciar la aplicación, se verifica si la base de datos existe.
   - Si no existe, se crean la base de datos y las tablas.
   - Si las tablas se crean, se inicializan con datos semilla.
3. **Interacción con la Base de Datos**:
   - El objeto `CN` permite ejecutar queries de forma persistente sin necesidad de reconexión.

## Notas Finales
Este proyecto integra conceptos avanzados de desarrollo en VB6, manejo de bases de datos y uso de herramientas externas para mejorar la calidad del código. Se recomienda seguir documentando futuras ampliaciones para mantener la profesionalización del repositorio.

**Repositorio:** Puedes encontrar el código y la documentación completa en GitHub (agregar enlace).


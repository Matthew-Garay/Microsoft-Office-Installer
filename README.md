# Microsoft Office Installer

## A. Titulo del Proyecto y Resumen Ejecutivo

Microsoft Office Installer es un conjunto de scripts de PowerShell que automatizan la instalacion de Microsoft Office en sistemas Windows mediante el Office Deployment Tool (ODT) de Microsoft. El problema que resuelve es el despliegue rapido y reproducible de Office sin instaladores manuales, ofreciendo dos modos de ejecucion: interfaz grafica y linea de comandos. Su alcance tecnologico se limita a sistemas Windows 10 y 11 con PowerShell 5.0 o superior, sin dependencias externas y sin gestion de licencias.

## B. Arquitectura del Sistema

El programa sigue una arquitectura de scripts procedimentales en PowerShell. El flujo de datos comienza con la seleccion de parametros por parte del usuario (edicion, idioma, aplicaciones, arquitectura), seguido de la descarga del Office Deployment Tool desde los servidores de Microsoft, la generacion de un archivo de configuracion XML con las opciones elegidas y la ejecucion de setup.exe con el parametro /configure. Existen dos puntos de entrada independientes: un script con interfaz grafica (instalar-gui.ps1) y un script de consola (instalar-cli.ps1), ambos siguiendo la misma secuencia logica. No existe arquitectura cliente-servidor, ni capas logicas separadas, ni persistencia de datos mas alla de los archivos temporales generados durante la instalacion.

## C. Caracteristicas Tecnicas Clave

- Dos modos de ejecucion: interfaz grafica y linea de comandos.
- Ediciones soportadas: LTSC Professional Plus 2024 y 2021; Professional Plus 2019, 2016 y 2013.
- Arquitectura de instalacion: 64 bits o 32 bits.
- Complementos opcionales: Project y Visio.
- Soporte para 10 idiomas: ingles, espanol, frances, aleman, portugues de Brasil, italiano, holandés, polaco, ruso y japones.
- Flujo guiado con confirmacion de cada opcion antes de modificar el sistema.
- Tema claro y oscuro en la version GUI.
- Sin dependencias externas y sin gestion de licencias (no incluye activacion).
- Registro de cada ejecucion en archivos de log temporales.
- No utiliza base de datos, ni concurrencia, ni procesamiento en tiempo real, ni protocolos de red propios.
- Seguridad basada en la confirmacion de usuario antes de realizar cambios en el sistema.

## D. Especificacion del Stack Tecnologico

| Categoria | Tecnologia / Libreria |
|-----------|----------------------|
| Lenguaje | PowerShell 5.0+ |
| Herramientas | Office Deployment Tool (ODT) de Microsoft |
| Plataforma | Windows 10 / 11 |
| Dependencias | Ninguna (sin dependencias externas) |

## E. Guia de Instalacion, Configuracion y Despliegue

### Requisitos previos del sistema

- Windows 10 u 11.
- PowerShell 5.0 o superior.
- Permisos de administrador.
- .NET Framework 4.5 o superior.
- Conexion a Internet (el ODT y los archivos de Office se descargan de Microsoft).

### Secuencia de comandos para clonacion e instalacion

git clone https://github.com/Matthew-Garay/Microsoft-Office-Installer.git
cd Microsoft-Office-Installer

### Configuracion de variables de entorno

Este proyecto no requiere archivo .env ni variables de entorno.

### Comandos para ejecucion en entorno de Desarrollo

Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass

Para la version con interfaz grafica:

irm https://raw.githubusercontent.com/Matthew-Garay/Microsoft-Office-Installer/main/instalar-gui.ps1 | iex

Para la version de consola:

irm https://raw.githubusercontent.com/Matthew-Garay/Microsoft-Office-Installer/main/instalar-cli.ps1 | iex

### Comandos para ejecucion en entorno de Produccion

No aplica entorno de produccion separado; los scripts se ejecutan localmente con privilegios de administrador.

## F. Creditos de Autoría

Desarrollado por: Matthew Ilveneff Garay Perez
Correo Electronico: 1911.garay.perez@gmail.com

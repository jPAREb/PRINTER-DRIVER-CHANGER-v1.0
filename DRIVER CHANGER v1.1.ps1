#################################
#       JORDI PARÉ BERNADÓ      #
#       |RICOH SANT CUGAT|      #
#################################
#        TWITTER ~@_xJPBx_~     #
#        INSTAGRAM ~jpareb~     #
# GMAIL~jparebernado@gmail.com~ #
#################################

$todas_impresoras_nombre = Get-Printer
$todas_impresoras_driver = Get-Printer
$contenidor_script = Read-Host -Prompt 'Carpeta donde se generarán los archivos'
$contenidor_script = $contenidor_script + "\Script_Cambiar_Drivers"
$contenidor_driver = $contenidor_script + "\DRIVER\"
$contenidor_logs = $contenidor_script + "\LOGS\"
$contenidor_printers = $contenidor_script + "\PRINTERS\"
$contenidor_printers_antiguas = $contenidor_script + "\IMPRESORAS ANTIGUAS\"
$lista_contenidores = @($contenidor_script, $contenidor_script, $contenidor_driver, $contenidor_logs, $contenidor_printers,$contenidor_printers_antiguas)
$archivo_log = $contenidor_logs + "LOG.TXT"
$archivo_impresoras = $contenidor_printers + "LISTADO IMPRESORAS.csv"
$archivo_impresoras_usuario = $contenidor_printers + "SELECCIÓN IMPRESORAS.csv"
$lista_archivos = @($archivo_log, $archivo_impresoras, $archivo_impresoras_usuario)
$contenedor_disk = $contenidor_driver + "disk1\"
$archivo_oemsetup = $contenedor_disk + "oemsetup.inf"
$nombre_driver = "RICOH PCL6 UniversalDriver V4.266"


function Apagar-Script {
    <#FUNCIÓN QUE DEBE SER LLAMADA PARA CERRAR EL SCRIPT, Y A LA VEZ, DAR UN PIE DE PÁGINA AL LOG#>
    Write-Host "ALGO HA FALLADO... REVISIÓN AL LOG"
    Generar-Log "¡¡EN BREVES SE CIERRA EL SCRIPT!!"
    Generar-Log "-----"
    Generar-Log "-FIN-"
    Generar-Log "-----"
    exit
}

function Generar-Log {
    <#FUNCIÓN QUE AL SER LLAMADA DEBE RECIBIR UN STRING PARA GUARDAR LOS MENSAJES DE LOS EVENTOS
    DE ESTA FORMA SE PUEDEN GENERAR LOS REGISTROS CON UN FORMATO#>
    param(
       [String] $valor1
    )
	if (Test-Path -Path $archivo_log){
    $fecha_hora = Get-Date -Format "[dd.MM.yyyy|HH.mm]"
    $mensaje_final = $fecha_hora + " " + $valor1
    write-output $mensaje_final | add-content $archivo_log
	}
}


function Menu-Impresoras{
    <#LA PRIMERA VEZ QUE SE LLAMA REGISTRA EN EL LOG UN "TÍTULO", DESDE AQUI SE PUEDE ADMINISTRAR
    EL ORDEN DE LOS PROCESOS#>
    Generar-Log "--------"
    Generar-Log "-INICIO-"
    Generar-Log "--------"
    Creacion-Directorios
    Creacion-Archivos
    Enviar-Impresoras
    $lista_impresora = Recibir-Impresora
    Driver-Check
    Generar-log "Se detectó correctamente el driver"
    $nombre_driver = Nombre-Driver
    $log = "Nombre extraído en el archivo oemsetup.inf: " + $nombre_driver
    Generar-Log $log
    Driver-Instalacion
    Driver-Impresoras $lista_impresora
    Generar-Log "SCRIPT FINALIZADO CORRECTAMENTE"
    Write-Host "FINALIZADO"
    Generar-Log "¡¡EN BREVES SE CIERRA EL SCRIPT!!"
    Generar-Log "-----"
    Generar-Log "-FIN-"
    Generar-Log "-----"
    exit
}




function Creacion-Directorios{
    <#LA FUNCIÓN GENERA AQUELLAS CARPETAS QUE NO LO ESTÁN#>
    Generar-Log "Se inicia el proceso de comprobación del estado de las carpetas"
    $i = 0
    $lista_contenidores | ForEach-Object{
        if(-not ($lista_contenidores[$i]|Test-Path)){
            $log = "No se ha generado la siguiente carpeta: " + $lista_contenidores[$i]
            Generar-Log $log
            $log = "Generando la carpeta: " + $lista_contenidores[$i]
            Generar-Log $log
            New-Item -Path $lista_contenidores[$i] -ItemType Directory
            
        }
        $log = "Ya ha sido generada la carpeta: " + $lista_contenidores[$i]
        Generar-Log $log
    $i++
    }
    Generar-Log "Directorios correctamente creados"
}

function Creacion-Archivos{
    <#LA FUNCIÓN GENERA AQUELLOS ARCHIVOS QUE NO LO ESTÁN#>
    Generar-Log "Se inicia el proceso de comprobación del estado de las archivos"
    $i = 0
    $lista_archivos | ForEach-Object{
        if(-not ($lista_archivos[$i]|Test-Path)){
            $log = "No se ha generado el siguiente archivo: " + $lista_archivos[$i]
            Generar-Log $log
            $log = "Generando el archivo: " + $lista_archivos[$i]
            Generar-Log $log
            New-Item -Path $lista_archivos[$i] -ItemType File
            if($archivo_impresoras_usuario -eq $lista_archivos[$i]){
                Add-Content -Path $lista_archivos[$i] -Value 'IMPRESORAS'
            }
        } elseif($archivo_impresoras -eq $lista_archivos[$i]){
            ListaImpresoras-Actualizar
        }
        $i++
    }
    Generar-Log "Archivos correctamente creados"
}

function ListaImpresoras-Actualizar {

			$nuevo_nombre = Get-Date -Format "dd.MM.yyyy-HH.mm.ss"
            $nuevo_nombre = $nuevo_nombre + ".csv"
            Rename-Item -Path $archivo_impresoras -NewName $nuevo_nombre
            $moviendo_archivo = $contenidor_printers + $nuevo_nombre
            Move-Item -Path  $moviendo_archivo -Destination $contenidor_printers_antiguas
            $log = "El siguiente archivo: " + $archivo_impresoras + " se mueve al siguiente destino: " + $contenidor_printers_antiguas + " con este nombre: " + $nuevo_nombre
            Generar-Log $log


}

function Enviar-Impresoras{
    <#TODAS LAS IMPRESORAS INSTALADAS EN EL SISTEMA, SERÁN REGISTRADAS EN EL ARCHIVO CORRESPONDIENTE#>
    Generar-Log "Se inicia el proceso para ver todas las impresoras"
    $i = 0
    Generar-Log "Listando impresoras del servidor"
    $todas_impresoras_nombre | ForEach-Object{
        $nombre = $todas_impresoras_nombre[$i].Name
        $driver = $todas_impresoras_nombre[$i].DriverName
        $final = $nombre + "," + $driver
        write-output $final | add-content $archivo_impresoras
        $i ++
    }
    $log = "Impresoras listadas en la siguiente ruta: " + $archivo_impresoras
    Generar-Log $log
}


function Recibir-Impresora{
    <#IMPORTA EN UNA ARRAY EL CONJUNTO DE IMPRESORAS#>
    $log = "Se inicia el proceso para leer las impresoras escritas para el usuario en el siguiente directorio: " + $archivo_impresoras_usuario
    Generar-Log $log
    $listado_impresoras_usuario = Import-Csv -Path $archivo_impresoras_usuario
    if($listado_impresoras_usuario){
        Generar-Log "Se han importado correctamente las impresoras escritas por el usuario"
    }else{
        Generar-Log "No se ha detectado ninguna impresora en la importación"
        $log = "Se requiere la comprobación en el siguiente archivo: " + $archivo_impresoras_usuario
        Generar-Log $log
        Apagar-Script
    }
    return $listado_impresoras_usuario
}

function Driver-Check{
    <#PARA EVITAR CRASHEOS, SE DETECTA SI EL ARCHIVO OEMSETUP.INF ESTÁ EN LA CARPETA DISK1 DE DRIVERS, EN CASO CONTRARIO
    SE PROCEDE A VER SI EXISTE EL ARCHIVO DISK1, QUEDA TODO REGISTRADO EN EL LOG#>
    $log = "Se inicia el proceso para comprobar si el driver está en el siguiente directorio: " + $contenidor_driver
    Generar-Log $log
    if(Test-Path -Path $archivo_oemsetup -PathType Any){
        Generar-Log "Archivo oemsetup encontrado con éxito"
    }else{
        $log = "No se encuentra el archivo oemsetup.inf en la siguiente ruta: " + $archivo_oemsetup
        Generar-Log $log
        $log = "Se procede a revisar si existe el contenedor: " + $contenedor_disk
        Generar-Log $log
        if(Test-Path -Path $contenedor_disk -PathType Any){
            Generar-Log "El contenedor existe"
        }else{
            Generar-Log "El contenedor no existe"
            $log = "Porfavor descarge el Driver que corresponda y deje la carpeta Disk1 y misc dentro de: " + $contenidor_driver
            Generar-Log $log
            Apagar-Script
        }
        Apagar-Script
    }
}

function Driver-Instalacion{
    <#LA FUNCION AL SER LLAMADA, LEE LA VARIABLE PÚBILCA CON EL ARCHIVO OEMSETUP.INF#>
    Generar-Log "Inicia el proceso de instalación"
    $llista = Get-Content -path $archivo_oemsetup   
    Invoke-Command {pnputil.exe -i -a $archivo_oemsetup}
    Add-PrinterDriver -Name $nombre_driver
    Generar-Log "El driver se ha instalado correctamente"
}


function Driver-Impresoras ($impresoras){
    <#LA FUNCION ESPECIFICA A LAS IMPRESORAS QUE DRIVER DEBEN USAR, TODO QUEDA REGISTRADO EN EL LOG#>
    $i = 0
    $impresoras | ForEach-Object{
       $log = "La impresora [" + $impresoras[$i].IMPRESORAS + "] se le instalará el driver"
       Generar-Log $log
       rundll32 printui.dll PrintUIEntry /Xs /n $impresoras[$i].IMPRESORAS DriverName $nombre_driver
       Generar-Log "Instalación completada con éxito"
       $i ++
    }
}


function Nombre-Driver{
    <#LA FUNCION BUSCA EL STRING 'CoDrvName' EN EL ARCHIVO OEMSETUP PARA ENCONTRAR EL NOMBRE DEL DIRVER#>
    $i = 0
    foreach($line in Get-Content $archivo_oemsetup) {
        $log = "Linia "+ $i + " del archivo oemsetup.inf"
        Generar-Log $log
        if($line -match 'CoDrvName ='){
            Generar-Log "Nombre del driver encontrado"
            $nombre_driver = $line -split """"
            return $nombre_driver[1]
        }
        $i ++
    } 
}

Menu-Impresoras

#################################
#       JORDI PARÉ BERNADÓ      #
#       |RICOH SANT CUGAT|      #
#################################
#        TWITTER ~@_xJPBx_~     #
#        INSTAGRAM ~jpareb~     #
# GMAIL~jparebernado@gmail.com~ #
#################################
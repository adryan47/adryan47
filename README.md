- 👋 Hi, I’m @adryan47

<!---
adryan47/adryan47 is a ✨ special ✨ repository because its `README.md` (this file) appears on your GitHub profile.
You can click the Preview link to take a look at your changes.
---> Hola estoy tratando de desplegar fuentes a traves de Script de Powershell, y no comprendo porque no instala las fuentes en C:\Windows\Font

#Definicion de Variables

#$NetworkPath = \\Software\Fonts\Gotham 

#Se define la variable $LocalPath1 la cual es una ubicacion local oculta en el perfil publico que por defecto el usuario no ve.
$LocalPath1 = "C:\FontsToDeploy\Fonts\*.*"

#Se define la variable LocalPath (lugar temporal para guardar fuentes)
$LocalPath = "C:\Users\Public\Fonts\"

$FONTS = 0x14


#Crea un directorio en $LocalPath
New-Item $LocalPath -type directory -Force

#Copia el contenido de la primer carpeta en la segunda
Copy-Item -Path "C:\FontsToDeploy\Fonts\*.*" -Destination "C:\Users\Public\Fonts\" -Force

#Ejcuta dir sobre la variable $LocalPath y guarda el resultado en la variable $Fontdir
$Fontdir = dir $LocalPath


#Notas-`New-Object` provides the most commonly-used functionality of the VBScript CreateObject function. 
#A statement like `Set objShell = CreateObject("Shell.Application")` in VBScript can be   translated 
#to `$objShell = New-Object -COMObject "Shell.Application"` in PowerShell. 

$Objshell = New-Object -COMObject "Shell.Application"

$shell = New-Object -ComObject "WScript.shell"

$objFolder = $objShell.Namespace($FONTS)

foreach($File in $Fontdir) 
{
  if (!(Test-Path "C:\Windows\Fonts\$File") -eq $False)

    {
     if(($File.name -match "Gotham*.ttf") -or ($File.name -match "Gotham*.otf"))
     {
            
         $objFolder.CopyHere($File.fullname)
        
            $shell.Namespace($LocalPath).Items().InvokeVerbEx("Install")
         #$File.InvokeVerb("Install")
    }
}
}

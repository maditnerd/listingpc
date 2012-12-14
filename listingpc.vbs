'Listing informatique
'Par Kelav
'Domaine Publique

'Adresses IP (variable: addresse_ip)  '
'-------------------------------------'      
   strComputer = "."
     
    Set objWMIService = GetObject("winmgmts:" & "!\\" & strComputer & "\root\cimv2")
    Set colAdapters = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled = True")
    For Each objAdapter in colAdapters
     
    IPdebut = LBound(objAdapter.IPAddress)
    IPfin = UBound(objAdapter.IPAddress)
    If (objAdapter.IPAddress(IPdebut) <> "") then
     
     
    For i = IPdebut To IPfin
	If InStr(objAdapter.IPAddress(i),":") = 0 Then 
	
	If objAdapter.IPAddress(i) = last_card Then
	Else
	'WScript.Echo objAdapter.IPAddress(i)
	'WScript.Echo last_card
	new_ip = objAdapter.IPAddress(i)
	if new_ip = "0.0.0.0" then new_ip = "Deconnecte"
	
	adresse_ip = adresse_ip & new_ip & ","

	
	
	End If
	
	End If
	last_card = objAdapter.IPAddress(i)
	Next
    
    End If
    Next
     
    adresse_ip = Left(adresse_ip,Len(adresse_ip)-1) 
    'Wscript.Echo adresse_ip

'Adresse Mac (variable: adresse_mac)
'----------------------------------' 
	
    strComputer = "."
     
    Set objWMIService = GetObject("winmgmts:" & "!\\" & strComputer & "\root\cimv2")
    Set colAdapters = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled = True")
    For Each objAdapter in colAdapters
     
    IPdebut = LBound(objAdapter.IPAddress)
    IPfin = UBound(objAdapter.IPAddress)
    If (objAdapter.IPAddress(IPdebut) <> "") then
     
    For i = IPdebut To IPfin
	If InStr(objAdapter.IPAddress(i),":") = 0 Then 

	If objAdapter.IPAddress(i) = last_cardmac Then
	Else
	'WScript.Echo objAdapter.MACAddress(i)
	'WScript.Echo last_cardmac
	adresse_mac = adresse_mac & objAdapter.MACAddress(i) & ","

	
	End If
	
	End If
	last_cardmac = objAdapter.IPAddress(i)
	Next
    
    End If
    Next 
     adresse_mac = Left(adresse_mac,Len(adresse_mac)-1) 
    'Wscript.Echo adresse_mac

'Version de l'OS (variable: systeme) 
'----------------------------------' 
	
strComputer = "."
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set oss = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")
For Each os in oss
systeme = os.caption
Next
'Wscript.Echo systeme


'Imprimantes connectés (variable: imprimantes)
'--------------------------------------------' 

Const ForAppending = 8 
Const ForReading = 1 

Dim WshNetwork, objPrinter, intDrive, intNetLetter

strComputer = "."

Set WshNetwork = CreateObject("WScript.Network") 
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
Set colInstalledPrinters = objWMIService.ExecQuery("Select * from Win32_Printer") 
Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystem",,48) 
Set WshShell = WScript.CreateObject("WScript.Shell") 
Set objFSO = CreateObject("Scripting.FileSystemObject") 

Set objPrinter = WshNetwork.EnumPrinterConnections
If objPrinter.Count = 0 Then
printer = "Pas d'imprimantes connectés "
else
For intDrive = 0 To (objPrinter.Count -1) Step 2
intNetLetter = IntNetLetter +1
if objPrinter.Item(intDrive +1) = "Microsoft XPS Document Writer" or objPrinter.Item(intDrive +1) = "Fax" Then
Else
printer = printer & objPrinter.Item(intDrive +1) & ","
End if

Next
end if
'Wscript.Echo printer

'Taille de mes documents variable: taille_mesdocuments
'Taille du bureau variable: taille_bureau
'-----------------------------------------------------
Dim fso
Set fso = WScript.CreateObject("Scripting.FileSystemObject")

Const MY_DOCUMENTS = &H5&

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Shell.Application")

Set objFolder = objShell.Namespace(MY_DOCUMENTS)
Set objFolderItem = objFolder.Self
strdocuments = objFolderItem.Path

taille_mesdocuments = getFolderSize(strdocuments)
taille_mesdocuments = taille_mesdocuments / 1048576
taille_mesdocuments = round(taille_mesdocuments,0)
taille_mesdocuments = taille_mesdocuments

set WshShell = WScript.CreateObject("WScript.Shell")
strbureau = WshShell.SpecialFolders("Desktop")

taille_bureau = getFolderSize(strbureau)
taille_bureau = taille_bureau / 1048576
taille_bureau = round(taille_bureau,0)
taille_bureau = taille_bureau


'WScript.Echo taille_mesdocuments
'WScript.Echo taille_bureau


Function getFolderSize(folderName)
    On Error Resume Next

    Dim folder
    Dim subfolder
    Dim size
    Dim hasSubfolders

    size = 0
    hasSubfolders = False

    Set folder = fso.GetFolder(folderName)
    ' Try the non-recursive way first (potentially faster?)
    Err.Clear
    size = folder.Size
    If Err.Number <> 0 then     ' Did not work; do recursive way:
        For Each subfolder in folder.SubFolders
            size = size + getFolderSize(subfolder.Path)
            hasSubfolders = True
        Next

        If not hasSubfolders then
            size = folder.Size
        End If
    End If

    getFolderSize = size

    Set folder = Nothing        ' Just in case
	

End Function


'Nom de l'ordinateur Variable: nom_ordinateur
'-------------------------------
Set wshNetwork = WScript.CreateObject( "WScript.Network" )
nom_ordinateur = wshNetwork.ComputerName
'WScript.Echo nom_ordinateur



'Ecriture dans le fichier texte

	'Recuperation du dossier du script

	dossier_du_script = left(WScript.ScriptFullName,(Len(WScript.ScriptFullName))-(len(WScript.ScriptName)))
	'WScript.Echo dossier_du_script
    

    nom_du_fichier = dossier_du_script & "Listing.txt"
    'WScript.Echo nom_du_fichier
if fso.FileExists(nom_du_fichier) Then
exists = true
Else
exists = false
End If	
	
	Const ForWriting = 8
    Set f = fso.OpenTextFile(nom_du_fichier, ForWriting,true)
	
	
	'Le script écrit dans le fichier texte ICI
	
	'VbCrLf (retour chariot)
	'nom_ordinateur = Nom de l'ordinateur (Netbios)
	'systeme = Version de Windows
	'adresse_ip = Adresse IP
	'adresse_mac = Adresse MAC
	'printer = Imprimantes connectés
	'taille_mesdocuments = Taille de mes documents
	'taille_bureau = Taille du bureau

if exists = false Then
f.write("Nom;Système d'exploitation;Adresse IP;Adresse MAC;Imprimantes;Taille Mes Documents (Mo);Taille Bureau (Mo)" & VbCrLf)	
End If

f.write(nom_ordinateur & ";" & systeme & ";" & adresse_ip & ";" & adresse_mac &  ";" & printer & ";" & taille_mesdocuments & ";" &  taille_bureau & ";" & VbCrLf)




Wscript.Quit
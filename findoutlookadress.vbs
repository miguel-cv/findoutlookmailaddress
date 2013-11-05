Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003

on error resume next

dim numerodesid
dim contador
dim strComputer
dim perfiles(10)
dim correos(10)
dim valor1
dim valor2

strComputer = "."

Rem Obtener lista de SID
sid=Obtenerlistasid()

rem ahora tenemos en numerodesid las sid que hay y en sid() los valores

Rem ahora buscamos los perfiles de correo que hay en cada SID

Obtenerperfiles(sid)

rem ahora miramos la cuenta de correo que hay

valor1=""
valor2=""

indice=0
For count = 1 to contador

rem obtenemos los valores de id

Obtenerid()

rem WScript.Echo "Valores ID Perfil"
rem WScript.echo "*" & valor1 & "*" & valor2 & "*"

procesa(valor1)
if valor1 <> valor2 then 
procesa (valor2)
end if
	
Next

rem Resultados

wscript.echo

temp=""

for count=0 to contador
  numero=len(correos(count))
  numero=numero-1
    if numero > 0 then
      temporal= left(correos(count),numero)
      correos(count)=temporal
      wscript.echo "Count:" & count
      wscript.echo "*" & correos (count) & "*"
    end if
next

wscript.quit

REM
REM OBTENERID

sub Obtenerid()
strComputer = "."
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_ 
strComputer & "\root\default:StdRegProv")
strKeyPath = perfiles (count) & "\9375CFF0413111d3B88A00104B2A6676"
strValueName = "{ED475418-B0D6-11D2-8C3B-00104B2A6676}"
oReg.GetBinaryValue HKEY_USERS,strKeyPath,strValueName,strValue
For i = LBound(strValue) To UBound (strValue)
  valor=valor & strValue (i)
Next
nuevo=StrReverse(valor)
valor2=Left(nuevo,4)
valor1=Right(nuevo,4)
end sub


rem PROCESA

sub procesa (ultimosdigitosid)
varValue=""
claveregistro="HKEY_USERS\" & perfiles (count) & "\9375CFF0413111d3B88A00104B2A6676\0000" & ultimosdigitosid & "\Email"
rem wscript.echo "CLAVE REGISTRO1:" & claveregistro
if perfiles(count) <> "" then
  Set Shell = CreateObject("WScript.Shell")
  arrValues  = Shell.RegRead(claveregistro)
  if not isnull (arrValues) then
    For intIndex = LBound(arrValues) To UBound(arrValues) Step 2
      varValue = varValue & Chr(arrValues(intIndex))
    Next
  arrValues = split(varValue, ";")
  For intIndex = LBound(arrValues) To UBound(arrValues)
    wscript.echo correos(indice) & arrValues (intIndex)
    correos(indice)=arrValues(intIndex)
    indice=indice+1
  Next
  end if
end if
end sub
     
REM OBTENER LISTA DE SID
         
Function Obtenerlistasid()
dim temp (10)
numerodesid=0
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery ("Select * from Win32_UserAccount Where LocalAccount = True")
For Each objItem in colItems
  if len (objItem.SID) > 10 then 
    numerodesid=numerodesid+1
    temp(numerodesid)=objItem.SID
  end if
Next
obtenerlistasid=temp
end function

REM OBTENER PERFILES

sub Obtenerperfiles(listadesid)
contador=0
rem wscript.echo "-"
for bucle = 1 to numerodesid
  Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
  strKeyPath = listadesid(bucle) & "\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles"
  oReg.EnumKey HKEY_USERS, strKeyPath, arrSubKeys
  if not isnull(arrSubKeys) then
    For Each subkey In arrSubKeys
      rem wscript.echo "SUBKEY IN ARRSUBKEYS:" & subkey & "*" 
      if subkey <> "" then
	contador=contador+1
	perfiles(contador)=strKeyPath & "\" & subkey
rem	wscript.echo "PERFIL (CONTADOR):" & perfiles(contador) 
rem wscript.echo "CONTADOR: " & contador
      end if 
    Next
  end if
next

end sub

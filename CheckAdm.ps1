#Chequea si ejecuto PW7 en modo admin
([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

#Si devuelve "True" estamos como Admin.
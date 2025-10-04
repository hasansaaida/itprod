
$serviceName = "FlaskAppService"
$batPath = "C:\ITPROD\start_flask_app.bat"

New-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name $serviceName -Value $batPath -PropertyType String -Force

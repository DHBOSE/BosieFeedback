# Stop the background server of this project (matches BosieFeedback webapp server.py only)
Get-CimInstance Win32_Process -Filter "Name='pythonw.exe'" |
  Where-Object { $_.CommandLine -like '*server.py*' -and $_.CommandLine -like '*BosieFeedback*' } |
  ForEach-Object { Stop-Process -Id $_.ProcessId -Force }

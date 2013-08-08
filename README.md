DeployMaster
============

A PowerShell script to assist SharePoint deployments. 

Provides a GUI and a very simple scripting engine, to automate deployment steps.

To run in GUI mode:
  .\DeployMaster.ps1
  
To run in automated script mode:
  .\DeployMaster.ps1 -NoGUI
  
To specify a different path and file name for the transcript:
  .\DeployMaster.ps1 -LogFilePath "C:\log.txt"
  .\DeployMaster.ps1 -LogFilePath "C:\log.txt" -NoGUI
  
NOTE: The config.xml file should always be present and placed in the same folder as the DeployMaster.ps1 file. Use the sample config.xml file as reference.

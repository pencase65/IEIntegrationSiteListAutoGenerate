@echo off
cd /d %~dp0
powershell (Get-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Edge').InternetExplorerIntegrationSiteList > InternetExplorerIntegrationSiteList.ini

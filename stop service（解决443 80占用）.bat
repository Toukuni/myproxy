@echo off
echo Stopping services...
net stop ReportServer /y
net stop MsDepSvc /y
net stop PeerDistSvc /y
net stop SyncShareSvc /y
net stop W3SVC /y
net stop WAS /y
net stop IISADMIN /y
net stop http /y
echo All services stopped.

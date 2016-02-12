for /L %%n in (1,1,1) do (
	
	ECHO ************* APP SUBMIT NUMBER: %%n *************
	c:\python27\python.exe %~dp0autosubmitWP.py %*
	c:\python27\python.exe %~dp0savescreenshotWP.py %*
	c:\python27\python.exe %~dp0uploadimages.py %*
	
)

::c:\python27\python.exe %~dp0autosubmitWP.py %*
::c:\python27\python.exe %~dp0savescreenshotWP.py %*
::c:\python27\python.exe %~dp0uploadimages.py %*
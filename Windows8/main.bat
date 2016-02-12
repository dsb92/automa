for /L %%n in (1,1,1) do (

	ECHO ************* APP SUBMIT NUMBER: %%n *************
	c:\python27\python.exe %~dp0autosubmitW8.py %*
)

::c:\python27\python.exe %~dp0buildappW8.py %*
::c:\python27\python.exe %~dp0savescreenshotW8.py %*
::c:\python27\python.exe %~dp0autosubmitW8.py %*

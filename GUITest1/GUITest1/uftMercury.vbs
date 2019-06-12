set qtp=createObject("quicktest.application")
qtp.visible=true
qtp.launch
qtp.open"D:\BOSTON DOCUMENTS\QA\Homework\GUITest1\GUITest1"
qtp.test.run
qtp.test.close
qtp.quit
set qtp=nothing
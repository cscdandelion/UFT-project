DataTable.ImportSheet "D:\BOSTON DOCUMENTS\QA\Homework\GUITest1\6255 UFT Project.xlsx",1,"Global"
n=DataTable.GetSheet("Global").GetRowCount

For i = 1 To n 
    Datatable.SetCurrentRow(i)
     
    firstName = DataTable.Value("FirstName", "Global")
    lastName = DataTable.Value("LastName","Global")
    phone = DataTable.Value("Phone", "Global")
    email = DataTable.Value("Email","Global")
    address = DataTable.Value("Address","Global")
    address2 = DataTable.Value("Address2","Global")
    city = DataTable.Value("City","Global")
    state = DataTable.Value("State","Global")
    postCode = DataTable.Value("PostCode","Global")
    userName = DataTable.Value("UserName","Global")
    password = DataTable.Value("Password","Global")
    confirmPassword = DataTable.Value("ConfirmPassword","Global")


Browser("chrome").Navigate "https://www.google.com/search?q=mercury+tours+website&oq=Mercury+tours+website&aqs=chrome.0.69i59j0l3.8527j1j8&sourceid=chrome&ie=UTF-8"
'wait(3)
Browser("chrome").Page("mercury tours website").Link("Mercury Tours - 974636").Click
wait(3)
Browser("chrome").Page("Welcome: Mercury Tours").Link("REGISTER").Click
Browser("chrome").Page("Register: Mercury Tours").CaptureBitmap "D:\BOSTON DOCUMENTS\QA\Homework\GUITest1\screenshot\" & timer & "End.png", True
Browser("chrome").Page("Register: Mercury Tours").WebEdit("firstName").Check CheckPoint("firstName") @@ script infofile_;_ZIP::ssf24.xml_;_
Browser("chrome").Page("Register: Mercury Tours").WebEdit("firstName").Set firstName @@ script infofile_;_ZIP::ssf6.xml_;_
Browser("chrome").Page("Register: Mercury Tours").WebEdit("lastName").Set lastName @@ script infofile_;_ZIP::ssf8.xml_;_
Browser("chrome").Page("Register: Mercury Tours").WebEdit("phone").Set phone @@ script infofile_;_ZIP::ssf9.xml_;_
Browser("chrome").Page("Register: Mercury Tours").WebEdit("userName").Set email @@ script infofile_;_ZIP::ssf11.xml_;_
Browser("chrome").Page("Register: Mercury Tours").WebEdit("address1").Set address @@ script infofile_;_ZIP::ssf13.xml_;_
Browser("chrome").Page("Register: Mercury Tours").WebEdit("address2").Set address2 @@ script infofile_;_ZIP::ssf14.xml_;_
Browser("chrome").Page("Register: Mercury Tours").WebEdit("city").Set city @@ script infofile_;_ZIP::ssf15.xml_;_
Browser("chrome").Page("Register: Mercury Tours").WebEdit("state").Set state @@ script infofile_;_ZIP::ssf16.xml_;_
Browser("chrome").Page("Register: Mercury Tours").WebEdit("postalCode").Set postCode @@ script infofile_;_ZIP::ssf17.xml_;_
Browser("chrome").Page("Register: Mercury Tours").WebEdit("email").Set userName @@ script infofile_;_ZIP::ssf18.xml_;_
Browser("chrome").Page("Register: Mercury Tours").WebEdit("password").SetSecure password @@ script infofile_;_ZIP::ssf19.xml_;_
Browser("chrome").Page("Register: Mercury Tours").WebEdit("confirmPassword").SetSecure confirmPassword @@ script infofile_;_ZIP::ssf21.xml_;_
Browser("chrome").Page("Register: Mercury Tours").CaptureBitmap "D:\BOSTON DOCUMENTS\QA\Homework\GUITest1\screenshot\" & timer & "End.png", True
Browser("chrome").Page("Register: Mercury Tours").Image("register").Check CheckPoint("register") @@ script infofile_;_ZIP::ssf23.xml_;_
Browser("chrome").Page("Register: Mercury Tours").Image("register").Click 43,3
Browser("chrome").Page("Register: Mercury Tours").CaptureBitmap "D:\BOSTON DOCUMENTS\QA\Homework\GUITest1\screenshot\" & timer & "End.png", True
wait(3)
Browser("chrome").Page("Register: Mercury Tours_2").Link("REGISTER").Click


Next
wait(3)

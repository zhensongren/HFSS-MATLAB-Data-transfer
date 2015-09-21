Option Explicit


   Dim oAnsoftApp

   Dim oDesktop

   Dim oProject

   Dim oDesign

   Dim oEditor

   Dim oModule

   Set oAnsoftApp = CreateObject("AnsoftHfss.HfssScriptInterface")

   Set oDesktop = oAnsoftApp.GetAppDesktop()

   Set oProject = oDesktop.GetActiveProject

   Set oDesign = oProject.GetActiveDesign()

   

'Matlab user settings
Dim omatlab
Dim Result
Dim matlab_visible 
Dim user_path, save_file_name 
matlab_visible = True 							'True to keep Matlab command window open, False to not display it
user_path = "C:\Program Files\MATLAB\R2008a"	'path that Matlab opens to and any data is saved to
save_file_name = "cst_data"						'name of .mat file matlab saves all data to

Dim nf
Dim f() 
Dim dd
Dim range
Dim j
Dim frequency


'Get thickness information
MsgBox("Please input the thickness of the slab")
dd = InputBox("Input the thickness of the slab in um","")
dd = CDbl(dd)


'Get an array of solved frequencies and convert to double
  set oModule = oDesign.GetModule("Solutions")
range = oModule.GetSolveRangeInfo("Setup1:Sweep1")
ReDim f(UBound(range))  
For j = LBound(range) To UBound(range) 
f(j)= CDbl(range(j))
Next
nf = UBound(range)+1


Dim MImag() 
Dim s11mag_mat()
Dim s11phase_mat()
Dim s21mag_mat()
Dim s21phase_mat()
Dim sendA(0)
Dim re_eps_eff() 
Dim im_eps_eff() 
Dim re_mu_eff() 
Dim im_mu_eff() 
Dim re_eps_av()
Dim im_eps_av()
Dim re_mu_av() 
Dim im_mu_av() 
Dim re_z_eff() 
Dim im_z_eff() 
Dim re_n_eff() 
Dim im_n_eff() 


ReDim s11mag_mat(nf-1) 
ReDim s11phase_mat(nf-1) 
ReDim s21mag_mat(nf-1)  
ReDim s21phase_mat(nf-1) 
ReDim re_eps_eff(nf-1) 
ReDim im_eps_eff(nf-1) 
ReDim re_mu_eff(nf-1) 
ReDim im_mu_eff(nf-1) 
ReDim re_eps_av(nf-1) 
ReDim im_eps_av(nf-1) 
ReDim re_mu_av(nf-1) 
ReDim im_mu_av(nf-1) 
ReDim re_z_eff(nf-1) 
ReDim im_z_eff(nf-1) 
ReDim re_n_eff(nf-1)
ReDim im_n_eff(nf-1)

'Creat output variables
Set oModule = oDesign.GetModule("OutputVariable")
oModule.CreateOutputVariable "mag_S11", "mag(S(WavePort1,WavePort1))",  _
  "Setup1 : Sweep1", "Modal Solution Data", Array()
oModule.CreateOutputVariable "rad_S11", "ang_rad(S(WavePort1,WavePort1))",  _
  "Setup1 : Sweep1", "Modal Solution Data", Array()
oModule.CreateOutputVariable "mag_S21", "mag(S(WavePort2,WavePort1))",  _
  "Setup1 : Sweep1", "Modal Solution Data", Array()
oModule.CreateOutputVariable "rad_S21", "ang_rad(S(WavePort2,WavePort1))",  _
  "Setup1 : Sweep1", "Modal Solution Data", Array()




'convert frequencies to GHz
ReDim frequency(nf-1)
For j = LBound(range) To UBound(range) 
frequency(j)=f(j)/1000000000
Next



'Get the values of the Output variables For the desired frequency.
For j = 0 To nf-1
   	s11mag_mat(j) = oModule.GetOutputVariableValue( "mag_S11", "Freq='" & frequency(j) & "GHz'",_
"Setup1 : Sweep1", _
"Modal Solution Data", _
Array())
   	s11phase_mat(j) = oModule.GetOutputVariableValue("rad_S11", "Freq='" & frequency(j) & "GHz'",_
"Setup1 : Sweep1", _
"Modal Solution Data", _
Array())
   	s21mag_mat(j) = oModule.GetOutputVariableValue("mag_S21","Freq='" & frequency(j) & "GHz'",_
"Setup1 : Sweep1", _
"Modal Solution Data", _
Array())
   	s21phase_mat(j) = oModule.GetOutputVariableValue("rad_S21","Freq='" & frequency(j) & "GHz'",_
"Setup1 : Sweep1", _
"Modal Solution Data", _
Array())
Next

' 
' Delete the output variables before finishing.
' 

oDesign.DeleteOutputVariable "mag_S11"
oDesign.DeleteOutputVariable "rad_S11"
oDesign.DeleteOutputVariable "mag_S21"
oDesign.DeleteOutputVariable "rad_S21"


'Matlab COM/ActiveX interaction
	'------------------------------------------------------------------------------------
	'1) create COM object and initiate Matlab as an activeX server
  
Set omatlab = CreateObject("Matlab.Desktop.Application")
If matlab_visible Then
    	Result = omatlab.Execute("h=actxserver('Matlab.Desktop.Application');")
    End If
    '2) send all MWS data to Matlab using Automation methods (see app note for listing)
Set omatlab = CreateObject("Matlab.Application")
         Call omatlab.PutWorkspaceData("s11mag","base",s11mag_mat)
	 Call omatlab.PutWorkspaceData("s11phase","base",s11phase_mat)
	 Call omatlab.PutWorkspaceData("s21mag","base",s21mag_mat)
	 Call omatlab.PutWorkspaceData("s21phase","base",s21phase_mat)
         Call omatlab.PutWorkspaceData("f","base",f)	
sendA(0) = dd/1000000

	 Call omatlab.PutWorkspaceData("A","base",sendA)
    '3) use Execute Command To control matlab engine
    'use .m files for ease of use, note that .m files must be in user_path or linked at startup!

    Result = omatlab.Execute("cd(user_path)")
    Result = omatlab.Execute("em_extraction_v4")

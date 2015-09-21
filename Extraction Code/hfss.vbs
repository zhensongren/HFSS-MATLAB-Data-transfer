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

Dim matlab_visible 
Dim user_path, save_file_name 
matlab_visible = True 							'True to keep Matlab command window open, False to not display it
user_path = "C:\Documents and Settings\padillaadmin\My Documents\MATLAB"	'path that Matlab opens to and any data is saved to
save_file_name = "cst_data"						'name of .mat file matlab saves all data to


'variable definitions
Dim nf
Dim f() 
Dim d
Dim range
Dim arr()
Dim j
MsgBox("Please input the thickness of the slab")

d = InputBox("Input the thickness of the slab in um","")

  set oModule = oDesign.GetModule("Solutions")

range = oModule.GetSolveRangeInfo("Setup1:Sweep1")





Dim MImag() 
Dim s11mag_mat()
Dim s11phase_mat()
Dim s22mag_mat()
Dim s22phase_mat()
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

ReDim f(nf-1)
ReDim s11mag_mat(nf-1)
ReDim s11phase_mat(nf-1)
ReDim s22mag_mat(nf-1)
ReDim s22phase_mat(nf-1)
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




'

' Get the values of the Output variables For the desired frequency.

'



For j = LBound(range) To UBound(range) 

   	f(j) = oDesign.GetOutputVariableValue( "Freq",_
"Setup1 : Sweep1", _
"Modal Solution Data", _
Array())

   	s11mag_mat(j) = oDesign.GetOutputVariableValue( "mag(S(WavePort1,WavePort1))", "Freq='f(j)'",_
"Setup1 : Sweep1", _
"Modal Solution Data", _
Array())

   	s11phase_mat(j) = oDesign.GetOutputVariableValue("ang_rad(S(WavePort1,WavePort1))", "Freq='f(j)'",_
"Setup1 : Sweep1", _
"Modal Solution Data", _
Array())
   	s21mag_mat(j) = oDesign.GetOutputVariableValue("mag(S(WavePort2,WavePort1))", "Freq='f(j)'",_
"Setup1 : Sweep1", _
"Modal Solution Data", _
Array())
   	s21phase_mat(j) = oDesign.GetOutputVariableValue("ang_rad(S(WavePort2,WavePort1))", "Freq='f(j)'",_
"Setup1 : Sweep1", _
"Modal Solution Data", _
Array())
Next 

'Matlab COM/ActiveX interaction

	'1) create COM object and initiate Matlab as an activeX server
    Set matlab = CreateObject("Matlab.Application")
    If matlab_visible Then
    	Result = matlab.Execute("h=actxserver('Matlab.Application');set(h,'visible',1);")
    End If
    '2) send all MWS data to Matlab using Automation methods (see app note for listing)
    Call matlab.PutCharArray("user_path","base",user_path)
    Call matlab.PutCharArray("save_file_name","base",save_file_name)
    Call matlab.PutFullMatrix("s11mag","base",s11mag_mat,MImag)
	Call matlab.PutFullMatrix("s11phase","base",s11phase_mat,MImag)
	Call matlab.PutFullMatrix("s22mag","base",s22mag_mat,MImag)
	Call matlab.PutFullMatrix("s22phase","base",s22phase_mat,MImag)
	Call matlab.PutFullMatrix("s21mag","base",s21mag_mat,MImag)
	Call matlab.PutFullMatrix("s21phase","base",s21phase_mat,MImag)
    Call matlab.PutFullMatrix("f","base",f,MImag)
		sendA(0) = d
	Call matlab.PutFullMatrix("sendA","base",sendA,MImag)

    '3) use Execute Command To control matlab engine
    'use .m files for ease of use, note that .m files must be in user_path or linked at startup!

    Result = matlab.Execute("cd(user_path);")
    Result = matlab.Execute("em_extraction_v4;")


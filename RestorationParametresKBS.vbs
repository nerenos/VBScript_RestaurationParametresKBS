Set fso = WScript.CreateObject("Scripting.FileSystemObject")
Dim repS : repS = "D:\KBW\APSDATA"
Dim repD : repD = "F:\tempdata"
Dim repO : repO = "F:\APSDATA"
Dim filA : filA = Array("Nproj.DB","Nproj.PX","Ntek.YG6","POS.XG0","POS.XG1","POS.XG2","POS.YG0","POS.YG1","POS.YG2","Pos1.DB","Pos1.PX","POS.DB","POS.PX","POS.XG3","POS.XG4","POS.XG5","POS.XG6","POS.YG3","POS.YG4","POS.YG5","POS.YG6","Proj.PX","Npos.DB","Profiel.DB","Profiel.PX","Proj.DB","Pos3.DB","Pos3.PX")
Dim route

If Fso.FolderExists(repD) Then
	Fso.DeleteFolder repD
End If

Set objFolder=fso.CreateFolder(repD)

For Each file in filA
	route = repS+"\"+file
	If Fso.FileExists(route) Then
		Fso.CopyFile route , repD+"\"+file
	Else
		MsgBox "Fichier "+route+" introuvable", vbExclamation
	End If
Next

If fso.FolderExists(repS) Then
	Fso.DeleteFolder repS
End If

Fso.copyfolder repO, repS

For Each file in filA
	route = repD+"\"+file
	If Fso.FileExists(route) Then
		Fso.CopyFile route , repS+"\"+file
	Else
		MsgBox "Fichier "+route+" introuvable", vbExclamation
	End If
Next
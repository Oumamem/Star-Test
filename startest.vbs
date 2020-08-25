Dim Con, Rec, MsG, A
Set Con = CreateObject("ADODB.Connection")
Set Rec = CreateObject("ADODB.Recordset")
 
On Error Resume Next
Con.Open "DRIVER={MySQL ODBC 5.3 ANSI Driver};SERVER=127.0.0.1;DATABASE=star;UID=root; port=3308"
If Err.Number <> 0 Then
    MsG = "Erreur N°" & Err.Number & vbCrLf _
    & "Description:" & vbCrLf & Err.Description & vbCrLf _
    & "Impossible d'ouvrire la BD"
    MsgBox MsG, vbCritical, "Erreur"
    Else
    A = "SELECT * FROM client"
    Rec.Open A, Con
    If Err.Number <> 0 Then
        Con.Close
        MsG = "Erreur N°" & Err.Number & vbCrLf _
        & "Description:" & vbCrLf & Err.Description & vbCrLf _
        & "Impossible d'ouvrire la table"
        MsgBox MsG, vbCritical, "Erreur"
        Else
        If Rec.EOF Then
            MsG = "Aucun enregistrement disponible pour cette requête"
            MsgBox MsG, vbInformation, ""
            Else 
            	Do While Not Rec.EOF 
           			 MsgBox Rec.Fields("code")
           			 Rec.MoveNext
           			 loop
        End If
        Con.Close
        Rec.Close
    End If
End If
Set Rec = Nothing
Set Con = Nothing
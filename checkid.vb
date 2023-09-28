Module checkid

    Public Function chkId(id) As Boolean
        Dim validate As Boolean = False
        Dim Century As Int16
        Dim year As Int16
        Dim month As Int16
        Dim day As Int16
        Dim Governorate As Int16

        If Len(id.ToString) <> 14 Then
            MsgBox("   الطول لا يساوى  14    " & Len(id.ToString))
            validate = False
            Return False
        Else
            validate = True
            Return True
        End If

        If IsNumeric(id) Then
            validate = True
            Return True
        Else
            MsgBox("مش رقم ")
            validate = False
            Return False
        End If

        If validate = True Then
            Century = Convert.ToInt16(Mid(id, 1, 1))
            year = Convert.ToInt16(Mid(id, 2, 2))
            month = Convert.ToInt16(Mid(id, 4, 2))
            day = Convert.ToInt16(Mid(id, 6, 2))
            Governorate = Convert.ToInt16(Mid(id, 8, 2))
        End If

        If (Century < 2 Or Century > 3) Or (month > 12) Or (day > 31) Or (Governorate > 27 And Governorate <> 88) Then
            MsgBox("الرقم القومى غير صحيح ", 2)
            Return False
        End If

    End Function
End Module

Sub Main()
    If ThisApplication.ActiveDocument Is Nothing Then
        MsgBox("Please open an assembly document!")
    ElseIf (ThisApplication.ActiveDocument.DocumentType <> kAssemblyDocumentObject) Then
        MsgBox("Please open an assembly document!")
    Else
    
        Dim oDoc As AssemblyDocument = ThisApplication.ActiveDocument
    
        GetInstancePropInfo(oDoc.ComponentDefinition.Occurrences)

    End If
End Sub

Sub GetInstancePropInfo(oOccus As ComponentOccurrences)

    Dim oOccu As ComponentOccurrence
    Dim oTempOccu As ComponentOccurrence
    
    ' The Instance Properties is accessiable via ComponentOccurrence only
    ' so below will get the ComponentOccurrence from ComponentOccurrenceProxy.
    For Each oTempOccu In oOccus
        If oTempOccu.Type = kComponentOccurrenceProxyObject Then
            oOccu = oTempOccu.NativeObject
             
        Else
            oOccu = oTempOccu
        End If
        
        MsgBox(oOccu.Name)
        '  Instance Properties
        If oOccu.OccurrencePropertySetsEnabled Then
            Dim oProp As Inventor.Property
            For Each oProp In oOccu.OccurrencePropertySets(1)
            
                ' Print property info
                MsgBox("    " & oProp.DisplayName & ":" & oProp.Expression)
            Next
        End If
        
        GetInstancePropInfo(oTempOccu.SubOccurrences)
    Next
End Sub
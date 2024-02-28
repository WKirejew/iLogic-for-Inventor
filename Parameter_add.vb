'Downloading variables 
Dim oParams As Parameters
Dim oApp As Inventor.Application = ThisApplication
'Checking if it's a part
Try
	Dim oPartDoc As PartDocument = ThisDoc.Document
	Dim oPartCompDef As PartComponentDefinition = oPartDoc.ComponentDefinition
	oParams = oPartCompDef.Parameters
Catch
	Exit Try
End Try
'Or an assembly:
Try
	Dim oAssyDoc As AssemblyDocument = ThisDoc.Document
	Dim oAssyCompDef As AssemblyComponentDefinition = oAssyDoc.ComponentDefinition
	oParams = oAssyCompDef.Parameters	
Catch
	Exit Try
End Try

Dim oUserParams As UserParameters = oParams.UserParameters

'Adding a text type parameter
Try
  p = Parameter("Index")
Catch
  oUserParams.AddByValue("Index","0", UnitsTypeEnum.kTextUnits)
End Try

'Adding an unitless parameter
Try
  p = Parameter("Utl")
Catch
  p = oUserParams.AddByExpression("Utl", "0", UnitsTypeEnum.kUnitlessUnits)
End Try

'Adding a distance type parameter
Try
  p = Parameter("dx_mm")
Catch
  p = oUserParams.AddByExpression("dx_mm", "10", UnitsTypeEnum.kMillimeterLengthUnits)
End Try

'Adding a boolean parameter
Try
  p = Parameter("Boolean")
Catch
  oUserParams.AddByValue("Boolean", False, UnitsTypeEnum.kBooleanUnits)
End Try

'More types: https://help.autodesk.com/view/INVNTOR/2023/ENU/?guid=GUID-59997AD8-527C-4552-B90E-88D1B1F97841
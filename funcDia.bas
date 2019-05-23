Attribute VB_Name = "funcDia"
Option Private Module
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///                                 Functions for Disparate Impact Analysis                                  ///
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///NOTES:                                                                                                    ///
'//Some of the Code in Funcs matrixDemoTotals & createSheetShell are redundant. Similarly, scope issues were  ///
'// resolved by declaring more variables than necessary. Code will be refactored in release v1.1.             ///
'//Function(s) for certain data validation features may be added in release v1.2.                             ///
'//Fishers Exact functions will be incorporated into release v2.0.                                            ///
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////

''Public Variables''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Min As Integer  'Mupltiple integers assigned per process related changes.         '
Public Minor As Integer  'First pass(0) for Major, Second pass(1) for Minor.             '
Public PctProgress As Double  'Declare param to update progress indicator.               '
Public SheetName As String  'Declare worksheet name string.                              '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Public Variables''

Public Function matrixDemoTotals(ByRef sAnalysisType As String, ByVal sAnalysisByValue As String) As Variant
'<Summary>
' sAnalysisType parameter indicates which Column the demographic data will be totaled by.
'  >1=Total for all records; 11=Total by Job Title; 12=Total by Department; 14=Total by Decision Maker
' sAnalysisValue indicates which AnalysisType value the demographic data will be totaled by.
'  >example: Totals for Employees with Job Title ScriptKitty || Totals for Employees in Department LaLaLand.
' Demographics (demo) value of TRUE increases the first set of integer variables declared, by 1 (Category variables).
'  >If Decision Column (2), 'has the employee been selected for RIF?' TRUE increase the RIF variables by 1.
'  >>FALSE & TRUE combo for NonCategory.
' Returns 5x4 Matrix of Minority, Female, AgeOver40, Disabled, Vet Fields/Columns Totaled By Category, NonCategory,
'  >CategoryRIF, NonCategoryRIF.
' OR returns 6x4 Matrix of H, A, BoAM, AIoAN, NHoOPI, ToMR Fields/Columns Totaled by Category, NonCategory,
'  >CategoryRIF, NonCategoryRIF.
' All Minority SubCategories are compared against the Non-Minority NonCategory (FALSE in Column 5 - Minority),
'  >rather than a FALSE value in the relevant Minority Subcategory demo column.
'<Summary>
    Dim iColumnNumber As Integer
    If Minor = 0 Then  'Running Major
        Dim matrix(1 To 5, 1 To 4) As Integer  'Variable assignment for return
        'Assign iColumnNumber based on param sAnalysisType
        Select Case sAnalysisType
            Case "All"
                iColumnNumber = 1
            Case "Title"
                iColumnNumber = 11
            Case "Dept"
                iColumnNumber = 12
            Case "DM"
                iColumnNumber = 14
        End Select
        
        'If params are outside acceptable range, show error message and exit the calling subroutine
        If (iColumnNumber = 1 And sValue2TotalBy <> Null) Or _
            (iColumnNumber = 11 And sValue2TotalBy = Null) Or _
            (iColumnNumber = 12 And sValue2TotalBy = Null) Or _
            (iColumnNumber = 14 And sValue2TotalBy = Null) Then
                MsgBox "Function totaling Demographics was called using unknown parameters. Process FAIL, EXIT Sub." _
                    , , "ERROR: unknwon matrixDemoTotals params"
                Exit Function 'ERROR: unknown matrixDemoTotals params
        End If
        
        'Category variables
        Dim iMinority, iFemale, iAgeOver40, iDisabled, iVet As Integer
            'Init Category variables
            iMinority = 0
            iFemale = 0
            iAgeOver40 = 0
            iDisabled = 0
            iVet = 0
        
        'Category RIF variables
        Dim iMinorityRif, iFemaleRif, iAgeOver40Rif, iDisabledRif, iVetRif As Integer
            'Init Category RIF variables
            iMinorityRif = 0
            iFemaleRif = 0
            iAgeOver40Rif = 0
            iDisabledRif = 0
            iVetRif = 0
        
        'NonCategory variables (iNonMinority Declared in Minor section)
        Dim iNonMinority, iMale, iAgeUnder40, iNotDisabled, iNonVet As Integer
            'Init NonCategory variables
            iNonMinority = 0
            iMale = 0
            iAgeUnder40 = 0
            iNotDisabled = 0
            iNonVet = 0
        
        'NonCategory RIF variables (iNonMinorityRif Declared in Minor section)
        Dim iNonMinorityRif, iMaleRif, iAgeUnder40Rif, iNotDisabledRif, iNonVetRif As Integer
            'Init NonCategory RIF variables
            iNonMinorityRif = 0
            iMaleRif = 0
            iAgeUnder40Rif = 0
            iNotDisabledRif = 0
            iNonVetRif = 0
                      
        'Condition change for sAnalysisType All vs other
        Dim condition1 As Boolean
        Dim condition2 As Boolean
        Dim condition3 As Boolean
        Dim condition4 As Boolean
        
        Dim nFrow As Long  'Loop Counter
        Dim lastRow As Long  'Last Row local assignment
            lastRow = nLastRow("Sheet1")
        Dim colmn As Integer
        For nFrow = 2 To lastRow
            For colmn = 6 To 10  'Loop through Columns(6) Minority through (10) Vet
                condition1 = False
                If sAnalysisType = "All" Then
                    'Category++ if demo value is True (e.g. Female? True)
                    If Sheets("Sheet1").Cells(nFrow, colmn).Value = True Then
                        condition1 = True
                    End If
                Else
                    'Category++ if demo true & AnalysisByValue match (e.g. Female ScriptKitty? True)
                    If Sheets("Sheet1").Cells(nFrow, colmn).Value = True And _
                        Sheets("Sheet1").Cells(nFrow, iColumnNumber).Value = sAnalysisByValue Then
                        condition1 = True
                    End If
                End If
                If condition1 = True Then
                    Select Case colmn
                        Case 6
                            iMinority = iMinority + 1
                        Case 7
                            iFemale = iFemale + 1
                        Case 8
                            iAgeOver40 = iAgeOver40 + 1
                        Case 9
                            iDisabled = iDisabled + 1
                        Case 10
                            iVet = iVet + 1
                    End Select
                End If
                
                condition2 = False
                If sAnalysisType = "All" Then
                    'NonCategory++ if demo value is False (e.g. Male? True)
                    If Sheets("Sheet1").Cells(nFrow, colmn).Value = False Then
                        condition2 = True
                    End If
                Else
                    'NonCategory++ if demo is False & AnalysisByValue match (e.g. Male ScriptKitty? True)
                    If Sheets("Sheet1").Cells(nFrow, colmn).Value = False And _
                        Sheets("Sheet1").Cells(nFrow, iColumnNumber).Value = sAnalysisByValue Then
                        condition2 = True
                    End If
                End If
                If condition2 = True Then
                    Select Case colmn
                        Case 6
                            iNonMinority = iNonMinority + 1
                        Case 7
                            iMale = iMale + 1
                        Case 8
                            iAgeUnder40 = iAgeUnder40 + 1
                        Case 9
                            iNotDisabled = iNotDisabled + 1
                        Case 10
                            iNonVet = iNonVet + 1
                    End Select
                End If
                
                condition3 = False
                If sAnalysisType = "All" Then
                    'CategoryRif++ if demo value is True and Decision value is True (e.g. Female RIF? True)
                    If Sheets("Sheet1").Cells(nFrow, colmn).Value = True And _
                        Sheets("Sheet1").Cells(nFrow, 2).Value = True Then
                        condition3 = True
                    End If
                Else
                    'CategoryRif++ if demo is True & sAnalysisByValue match & Decision value is True (e.g. Female ScriptKitty RIF? True)
                    If Sheets("Sheet1").Cells(nFrow, colmn).Value = True And _
                        Sheets("Sheet1").Cells(nFrow, iColumnNumber).Value = sAnalysisByValue And _
                            Sheets("Sheet1").Cells(nFrow, 2).Value = True Then
                        condition3 = True
                    End If
                End If
                If condition3 = True Then
                    Select Case colmn
                        Case 6
                            iMinorityRif = iMinorityRif + 1  'CategoryRIF++
                        Case 7
                            iFemaleRif = iFemaleRif + 1
                        Case 8
                            iAgeOver40Rif = iAgeOver40Rif + 1
                        Case 9
                            iDisabledRif = iDisabledRif + 1
                        Case 10
                            iVetRif = iVetRif + 1
                    End Select
                End If
                    
                condition4 = False
                If sAnalysisType = "All" Then
                    'NonCategoryRif++ if demo value is False and Decision value is True (e.g. Male RIF? True)
                    If Sheets("Sheet1").Cells(nFrow, colmn).Value = False And _
                        Sheets("Sheet1").Cells(nFrow, 2).Value = True Then
                        condition4 = True
                    End If
                Else
                    'NonCategoryRif++ if demo is False & sAnalysisValue match & Decision value is True (e.g. Male ScriptKitty RIF? True)
                    If Sheets("Sheet1").Cells(nFrow, colmn).Value = False And _
                        Sheets("Sheet1").Cells(nFrow, iColumnNumber).Value = sAnalysisByValue And _
                            Sheets("Sheet1").Cells(nFrow, 2).Value = True Then
                        condition4 = True
                    End If
                End If
                If condition4 = True Then
                    Select Case colmn
                        Case 6
                            iNonMinorityRif = iNonMinorityRif + 1
                        Case 7
                            iMaleRif = iMaleRif + 1
                        Case 8
                            iAgeUnder40Rif = iAgeUnder40Rif + 1
                        Case 9
                            iNotDisabledRif = iNotDisabledRif + 1
                        Case 10
                            iNonVetRif = iNonVetRif + 1
                    End Select
                End If
            Next
        Next
        
        
        '///Init 5x4 Matrix with totals///
        '/////In cell placement order/////
        'Age
        matrix(1, 1) = iAgeOver40
        matrix(1, 2) = iAgeUnder40
        matrix(1, 3) = iAgeOver40Rif
        matrix(1, 4) = iAgeUnder40Rif
        'Gender
        matrix(2, 1) = iFemale
        matrix(2, 2) = iMale
        matrix(2, 3) = iFemaleRif
        matrix(2, 4) = iMaleRif
        'Disability Status
        matrix(3, 1) = iDisabled
        matrix(3, 2) = iNotDisabled
        matrix(3, 3) = iDisabledRif
        matrix(3, 4) = iNotDisabledRif
        'Race/Ethnicity
        matrix(4, 1) = iMinority
        matrix(4, 2) = iNonMinority
        matrix(4, 3) = iMinorityRif
        matrix(4, 4) = iNonMinorityRif
        'Protected Veteran Status
        matrix(5, 1) = iVet
        matrix(5, 2) = iNonVet
        matrix(5, 3) = iVetRif
        matrix(5, 4) = iNonVetRif
    
        'Return assignment
        matrixDemoTotals = matrix
    Else
    '/Minor: START->
    ' *Change upper inner index from 5 to 6 if for Minority Subcategories.
    ' *Change variable names of Total values (unnecessary but more readable).
    ' *Rather than Totaling False values for NonCateogry,
    '   all Minority Subcategories are compared against NonMinority
    '   (False in Column 5 of Sheet1) values.
        Dim matrixMin(1 To 6, 1 To 4) As Integer  'Variable assignment for return
        'Assign iColumnNumber based on param sAnalysisType
        Select Case sAnalysisType
            Case "All"
                iColumnNumber = 1
            Case "Title"
                iColumnNumber = 11
            Case "Dept"
                iColumnNumber = 12
            Case "DM"
                iColumnNumber = 14
        End Select
        
        Dim iNHoOPI, iBoAM, iA, iAIoAN, iToMR, iH As Integer  'Category variables
            iNHoOPI = 0
            iBoAM = 0
            iA = 0
            iAIoAN = 0
            iToMR = 0
            iH = 0
        
        Dim iNHoOPIRif, iBoAMRif, iARif, iAIoANRif, iToMRRif, iHRif As Integer  'Category RIF variables
            iNHoOPIRif = 0
            iBoAMRif = 0
            iARif = 0
            iAIoANRif = 0
            iToMRRif = 0
            iHRif = 0
        
            'NonCategory variables - already Declared in Major
            iNonMinority = 0
            iNonMinorityRif = 0
        
        Dim fRow As Long  'Loop Counter
        Dim ltRow As Long  'Last Row Count Local assignment
            ltRow = nLastRow("Sheet1")
        Dim iFcolumn2 As Integer  'Loop Counter
        For fRow = 2 To ltRow
            For iFcolumn2 = 5 To 21  'Loop through Columns(6) Minority through (10) Vet
                'Skip over Columns between Hispanic(4) and Asian(17)
                If iFcolumn2 = 6 Or iFcolumn2 = 7 Or iFcolumn2 = 8 Or iFcolumn2 = 9 Or _
                    iFcolumn2 = 10 Or iFcolumn2 = 11 Or iFcolumn2 = 12 Or iFcolumn2 = 13 Or _
                        iFcolumn2 = 14 Or iFcolumn2 = 15 Or iFcolumn2 = 16 Then
                    GoTo SkipColumnsBetweenHandA
                End If
                
                Dim condition1min As Boolean
                condition1min = False
                If sAnalysisType = "All" Then
                    'Category++ if demo value is True (e.g. Hispanic? True)
                    If Sheets("Sheet1").Cells(fRow, iFcolumn2).Value = True Then
                        condition1min = True
                    End If
                Else
                    'Category++ if demo true & sAnalysisByValue match (e.g. Hispanic ScriptKitty? True)
                    If Sheets("Sheet1").Cells(fRow, iFcolumn2).Value = True And _
                        Sheets("Sheet1").Cells(fRow, iColumnNumber).Value = sAnalysisByValue Then
                        condition1min = True
                    End If
                End If
                    If condition1min = True Then
                        Select Case iFcolumn2
                            Case 5
                                iH = iH + 1
                            Case 17
                                iA = iA + 1
                            Case 18
                                iBoAM = iBoAM + 1
                            Case 19
                                iAIoAN = iAIoAN + 1
                            Case 20
                                iNHoOPI = iNHoOPI + 1
                            Case 21
                                iToMR = iToMR + 1
                        End Select
                    End If
            
                Dim condition2min As Boolean
                condition2min = False
                If sAnalysisType = "All" Then
                    'NonCategory++ if Minority value is False (e.g. White(NonMinority)? True)
                    If Sheets("Sheet1").Cells(fRow, 6).Value = False Then
                        condition2min = True
                    End If
                Else
                    'NonCategory++ if Minority is False & sAnalysisByValue match (e.g. White ScriptKitty? True)
                    If Sheets("Sheet1").Cells(fRow, 6).Value = False And _
                        Sheets("Sheet1").Cells(fRow, iColumnNumber).Value = sAnalysisByValue Then
                        condition2min = True
                    End If
                End If
                    If condition2min = True Then
                        Select Case iFcolumn2
                            Case 5
                                iNonMinority = iNonMinority + 1
                        End Select
                    End If
                
                Dim condition3min As Boolean
                condition3min = False
                If sAnalysisType = "All" Then
                    'CategoryRif++ if demo value is True and Decision value is True (e.g. Asian RIF? True)
                    If Sheets("Sheet1").Cells(fRow, iFcolumn2).Value = True And _
                        Sheets("Sheet1").Cells(fRow, 2).Value = True Then
                        condition3min = True
                    End If
                Else
                    'CategoryRif++ if demo value is True & AnalysisValue is True & Decision value is True (e.g. Asian ScriptKitty RIF? True)
                    If Sheets("Sheet1").Cells(fRow, iFcolumn2).Value = True And _
                        Sheets("Sheet1").Cells(fRow, iColumnNumber).Value = sAnalysisByValue And _
                            Sheets("Sheet1").Cells(fRow, 2).Value = True Then
                        condition3min = True
                    End If
                End If
                    If condition3min = True Then
                        Select Case iFcolumn2
                                Case 5
                                    iHRif = iHRif + 1
                                Case 17
                                    iARif = iARif + 1
                                Case 18
                                    iBoAMRif = iBoAMRif + 1
                                Case 19
                                    iAIoANRif = iAIoANRif + 1
                                Case 20
                                    iNHoOPIRif = iNHoOPIRif + 1
                                Case 21
                                    iToMRRif = iToMRRif + 1
                        End Select
                    End If
                
                Dim condition4min As Boolean
                condition4min = False
                If sAnalysisType = "All" Then
                    'NonCategoryRif++ if demo value is False and Decision value is True (e.g. Male RIF? True)
                    If Sheets("Sheet1").Cells(fRow, 6).Value = False And _
                        Sheets("Sheet1").Cells(fRow, 2).Value = True Then
                        condition4min = True
                    End If
                Else
                    'NonCategoryRif++ if demo is False & sAnalysisByValue match & Decision value is True (e.g. Male ScriptKitty RIF? True)
                    If Sheets("Sheet1").Cells(fRow, 6).Value = False And _
                        Sheets("Sheet1").Cells(fRow, iColumnNumber).Value = sAnalysisByValue And _
                            Sheets("Sheet1").Cells(fRow, 2).Value = True Then
                        condition4min = True
                    End If
                End If
                    If condition4min Then
                        Select Case iFcolumn2
                            Case 5
                                iNonMinorityRif = iNonMinorityRif + 1
                        End Select
                    End If
SkipColumnsBetweenHandA:
            Next
        Next

        '///Init 6x4 Matrix with totals///
        '/////In cell placement order/////
        'Asian
        matrixMin(1, 1) = iA
        matrixMin(1, 2) = iNonMinority
        matrixMin(1, 3) = iARif
        matrixMin(1, 4) = iNonMinorityRif
        'American Indian or Alaskan Native
        matrixMin(2, 1) = iAIoAN
        matrixMin(2, 2) = iNonMinority
        matrixMin(2, 3) = iAIoANRif
        matrixMin(2, 4) = iNonMinorityRif
        'Black or African American
        matrixMin(3, 1) = iBoAM
        matrixMin(3, 2) = iNonMinority
        matrixMin(3, 3) = iBoAMRif
        matrixMin(3, 4) = iNonMinorityRif
        'Hispanic
        matrixMin(4, 1) = iH
        matrixMin(4, 2) = iNonMinority
        matrixMin(4, 3) = iHRif
        matrixMin(4, 4) = iNonMinorityRif
        'Native Hawaiian or Other Pacific Islander
        matrixMin(5, 1) = iNHoOPI
        matrixMin(5, 2) = iNonMinority
        matrixMin(5, 3) = iNHoPIRif
        matrixMin(5, 4) = iNonMinorityRif
        'Two or More Races
        matrixMin(6, 1) = iToMR
        matrixMin(6, 2) = iNonMinority
        matrixMin(6, 3) = iToMRRif
        matrixMin(6, 4) = iNonMinorityRif
    
        'Return assignment
        matrixDemoTotals = matrixMin
    
    'Minor: <-END/
    End If
            
End Function

Public Function rDif(ByVal matrix As Variant) As Variant
'<Summary>
' Calculate the differece between expected and actual selections.
' Result is rounded to an integer as it represents an individual (no fractions of people).
' Returned array contains values for the 5 major demo categories.
' /In cell placement order/
'<Summary>

    '/Minor: Change index from 5 to 6 if for Minority Subcategories
        If Minor = 1 Then  'Running Minor
            Min = 6  '6 Minority SubCategories
        Else
            Min = 5  '5 Major Categories
        End If
    '\Minor: Change index from 5 to 6 if for Minority Subcategories
    
    Dim dif As Integer  'For return array value assignment
    Dim d() As Integer  'Variable assignment for return
        ReDim d(1 To Min)  'ReDim to size
    Dim index As Integer  'Loop Counter
    For index = 1 To Min  'Loop through demo categories (AgeOver40, Female, Disabled, Minority, Vet)
        On Error GoTo CantDivideby0a  'I'm sure there's a much easier way to do this... workaround to be refactored later
        dif = matrix(index, 3) - (Int(((matrix(index, 4) + matrix(index, 3)) / (matrix(index, 2) + matrix(index, 1))) * matrix(index, 1)))
            If index = -1 Then
CantDivideby0a:
                dif = 0
                d(index) = dif
                Resume Next
            End If
        d(index) = dif
    Next
    
    'Return assignment
    rDif = d
    
End Function

Public Function rImpactRatio(ByVal matrix As Variant) As Variant
'<Summary>
' Calculate impact ratio between expected and actual selection rates (80% rule).
' Result is returned as Double, formated in the calling sub to show 2 decimal places.
'  >Positive result > 0.8 indicates potential AI for the category (e.g. Female)
'  >Negative result < -0.8 indicates potential AI for the noncategory (e.g. Male)
' Returned array contains values for the 5 major demo categories.
' /In cell placement order/
'<Summary>

    '/Minor: Change index from 5 to 6 if for Minority Subcategories
        If Minor = 1 Then  'Running Minor
            Min = 6  '6 Minority SubCategories
        Else
            Min = 5  '5 Major Categories
        End If
    '\Minor: Change index from 5 to 6 if for Minority Subcategories

    Dim dIR As Variant  'For return array value assignment - double type unless dividing by 0, in which case empty string
    Dim rIR As Variant  'Variable assignment for return
        ReDim rIR(1 To Min)  'ReDim to size
    Dim index As Integer  'Loop Counter
    For index = 1 To Min  'Loop through demo categories (AgeOver40, Female, Disabled, Minority, Vet)
        On Error GoTo CantDivideby0b  'I'm sure there's a much easier way to do this... workaround to be refactored later
        dIR = (matrix(index, 4) / matrix(index, 2)) / (matrix(index, 3) / matrix(index, 1))
            If index = -1 Then
CantDivideby0b:
                dIR = ""
                rIR(index) = dIR
                Resume Next
            End If
        rIR(index) = Round(dIR, 2)
    Next
    
    'Return assignment
    rImpactRatio = rIR
    
End Function

Public Function rStandardDeviation(ByVal matrix As Variant) As Variant
'<Summary>
' Calculate standard deviation (applying Stand Deviation of 2.0).
' 95% confidence interval measure of whether the difference between expected and actual selection rates are due to chance.
' Result is returned as a Double, formated in the calling sub to show 2 decimal places.
'  >Result >2.0 indicates likelihood that difference may not be due to chance
' The value of result determines the strength of that indication.
'  >e.g. 2SD of 3.0 indicates greater probability that differences are not due to chance, than indicated by 2SD of 2.0
' /In cell placement order/
'<Summary>

    '/Minor: Change index from 5 to 6 if for Minority Subcategories
        If Minor = 1 Then  'Running Minor
            Min = 6  '6 Minority SubCategories
        Else
            Min = 5  '5 Major Categories
        End If
    '\Minor: Change index from 5 to 6 if for Minority Subcategories

    Dim dSD As Variant  'For return array value assignment - double type unless dividing by 0, in which case empty string
    Dim rSD As Variant  'Variable assignment for return
        ReDim rSD(1 To Min)  'ReDim to size
    Dim index As Integer  'Loop Counter
    For index = 1 To Min  'Loop through demo categories (AgeOver40, Female, Disabled, Minority, Vet)
        On Error GoTo CantDivideby0c  'I'm sure there's a much easier way to do this... workaround to be refactored later
        dSD = ((matrix(index, 3) / matrix(index, 1)) - (matrix(index, 4) / matrix(index, 2))) / _
            Sqr((((matrix(index, 4) + matrix(index, 3)) / (matrix(index, 1) + matrix(index, 2))) _
                * (1 - ((matrix(index, 4) + matrix(index, 3)) / (matrix(index, 1) + matrix(index, 2))))) / _
                    ((matrix(index, 1) + matrix(index, 2)) * (matrix(index, 1) / (matrix(index, 1) + _
                        matrix(index, 2))) * (1 - (matrix(index, 1) / (matrix(index, 2) + matrix(index, 1))))))
            If index = -1 Then
CantDivideby0c:
                dSD = ""
                rSD(index) = dSD
                Resume Next
            End If
        rSD(index) = Round(dSD, 2)
    Next
    
    'Return assignment
    rStandardDeviation = rSD
    
End Function

Public Function matrixImpactedEmployees(ByRef sAnalysisByValue As Variant, ByRef sAnalysisType As String, ByVal sImpactDemo As String) As Variant
'<Summary>
' If Adverse Impact is indicated for an Analysis Value (e.g. employees with Job Title ScriptKitty).
' Loop through Worksheet containing source data using the Analysis Type (e.g. 'Title', 'Dept', or 'Decision Maker') Column.
' If employee record's Analysis Value matches param,
'  >and employee record's Demo value for impacted category is True,
'  >>and employee has been selected for RIF (True in Column 2)
' Then add employee's ID, Fname, and Lname to returned matrix.
'<Summary>
    
    Dim matrix As Variant
    Dim iAnalysisByColumn As Integer  'Declare column of Analysis Type
    Select Case sAnalysisType  'Param assigned by subDia
        Case "All"
            iAnalysisByColumn = 1
        Case "Title"
            iAnalysisByColumn = 11
        Case "Dept"
            iAnalysisByColumn = 12
        Case "DM"
            iAnalysisByColumn = 14
    End Select
    
    Dim iImpactDemoColumn As Integer  'Declare column of ImpactDemo
    Select Case sImpactDemo  'Param passed as Cell.Value of Demo category assigned X in Adverse Impact Column
        Case "Hispanic"
            iImpactDemoColumn = 5
        Case "Minority"
            iImpactDemoColumn = 6
        Case "Female"
            iImpactDemoColumn = 7
        Case "Age Over 40"
            iImpactDemoColumn = 8
        Case "Disabled"
            iImpactDemoColumn = 9
        Case "Veteran"
            iImpactDemoColumn = 10
        Case "Asian"
            iImpactDemoColumn = 17
        Case "Black or African American"
            iImpactDemoColumn = 18
        Case "American Indian or Alaskan Native"
            iImpactDemoColumn = 19
        Case "Native Hawaiian or Other Pacific Islander"
            iImpactDemoColumn = 20
        Case "Two or More Races"
            iImpactDemoColumn = 21
    End Select
    
    Sheets("Sheet1").Activate
    
    Dim nFrow1 As Integer  'Loop Counter
    Dim index As Integer  'Declare index of returned array
        index = 0  'Init index
    lastRow = nLastRow("Sheet1")
    'First Loop through criteria to determine matrix size
    For nFrow1 = 2 To lastRow  'Loop through all rows of source data Worksheet (exclude header)
        'Param sAnalysisByValue passed as Cell.Value of Analysis Sub Pool (e.g. Analysis for all EEs within CSSAA Dept: Cell.Value = "CSSAA")
        '>Param in the same row as the Cell.Value of Demo category assigned X in Adverse Impact Column
        If Sheets("Sheet1").Cells(nFrow1, iAnalysisByColumn) Like sAnalysisByValue And _
            Sheets("Sheet1").Cells(nFrow1, iImpactDemoColumn) = True And _
            Sheets("Sheet1").Cells(nFrow1, 2) = True Then
                index = index + 1
        End If
    Next
    
    'ReDim matrix to correct size
    ReDim matrix(1 To index + 1, 1 To 3)
    index = 0  'Reassign index to initial value
    
    'Second Loop through same criteria to populate matrix
    For nFrow1 = 2 To lastRow  'Loop through all rows of source data Worksheet (exclude header)
        'Param sAnalysisByValue passed as Cell.Value of Analysis Sub Pool (e.g. Analysis for all EEs within CSSAA Dept: Cell.Value = "CSSAA")
        '>Param in the same row as the Cell.Value of Demo category assigned X in Adverse Impact Column
        If Sheets("Sheet1").Cells(nFrow1, iAnalysisByColumn) Like sAnalysisByValue And _
            Sheets("Sheet1").Cells(nFrow1, iImpactDemoColumn) = True And _
            Sheets("Sheet1").Cells(nFrow1, 2) = True Then
                index = index + 1
                matrix(index, 1) = Sheets("Sheet1").Cells(nFrow1, 1)  'EE ID
                matrix(index, 2) = Sheets("Sheet1").Cells(nFrow1, 3)  'EE Fname
                matrix(index, 3) = Sheets("Sheet1").Cells(nFrow1, 4)  'EE Lname
        End If
    Next
    
    'Return assignment
    matrixImpactedEmployees = matrix
    
End Function

Public Function rUniqueAnalysisValues(ByRef sAnalysisType As String) As Variant
'<Summary>
' Return an array containing just UniqueAnalysisValues of the AnalysisType.
' This array is used by subDia to format parts of the report.
' Each unique value is given five rows in the relevant "Analysis By" Worksheet
'  >for each of the five major demo categories.
'<Summary>

    Dim iAnalysisByColumn As Integer  'Declare column of Analysis By Value (e.g. Dept CSSAA)
    
    Select Case sAnalysisType  'Param assigned by subDia
        Case "Title"
            iAnalysisByColumn = 11
        Case "Dept"
            iAnalysisByColumn = 12
        Case "DM"
            iAnalysisByColumn = 14
    End Select
    
    If sAnalysisType <> "All" Then
        'Create a copy of the AnalysisByColumn in the first empty column in the Worksheet
        Sheets("Sheet1").Select
            Columns(iAnalysisByColumn).Select
                Selection.Copy
            Columns(22).Select
                ActiveSheet.Paste
            
        'Remove duplicates in the new column
        Dim nlrow As Long
        nlrow = nLastRow("Sheet1")
        Sheets("Sheet1").Range(Sheets("Sheet1").Cells(1, 22), Sheets("Sheet1").Cells(nlrow, 22)).RemoveDuplicates Columns:=1, Header:=xlYes
        
        Dim nFrow2 As Integer  'Loop Counter
        Dim lastRow As Integer
        lastRow = nLastRowofColumn("Sheet1", 22)
        
        'ReDim rValues to corrent array size
        Dim rValues As Variant
        ReDim rValues(1 To lastRow - 1) As Variant
        
        For nFrow2 = 2 To lastRow 'Loop through the rows new column contains only unique values
            If Not sAnalysisType = "All" Then
                rValues(nFrow2 - 1) = Sheets("Sheet1").Cells(nFrow2, 22).Value 'Assign each unique ID value to returned array
            End If
        Next
        
        'Return assignment
        rUniqueAnalysisValues = rValues
        
        'Delete the unique value column created at top of function
        Sheets("Sheet1").Columns(22).Delete
    Else
        'Temp assignment for return
        Dim temp As Variant
        ReDim temp(1 To 1)
        temp(1) = "*"
        
        'Return assignment
        rUniqueAnalysisValues = temp
    End If

End Function

Public Function CreateSheetshell(ByRef sAnalysisType As String)
'<Summary>
' Creates a "Shell" of the analysis worksheet
'  >analysis segment for each unique value of Analysis By field/column
'  >>segment consists of one line per each of the five major demo categories under analysis
' Headers for entire worksheet
'<Summary>
    
    '/Minor: Change index from 5 to 6 if for Minority Subcategories
        If Minor = 1 Then  'Running Minor
            Min = 6  '6 Minority SubCategories
        Else
            Min = 5  '5 Major Categories
        End If
    '\Minor: Change index from 5 to 6 if for Minority Subcategories
        
    Dim rString As Variant
    Dim iColumn As Integer  'Declare demo category labels column
    Dim sColumn As String  'Declare Analysis Value header
    
    If Minor = 1 Then
        Select Case sAnalysisType 'SheetName assignment using param
            Case "All"
                SheetName = "By All Considered (MinSub)"
                iColumn = 1
                sColumn = "None"
            Case "Title"
                SheetName = "By Job Title (MinSub)"
                iColumn = 2
                sColumn = "Job Title"
            Case "Dept"
                SheetName = "By Department (MinSub)"
                iColumn = 2
                sColumn = "Department"
            Case "DM"
                SheetName = "By Decision Maker (MinSub)"
                iColumn = 4
                sColumn = "Decision Maker"
        End Select
    Else
        Select Case sAnalysisType 'SheetName assignment using param
            Case "All"
                SheetName = "By All Considered"
                iColumn = 1
                sColumn = "None"
            Case "Title"
                SheetName = "By Job Title"
                iColumn = 2
                sColumn = "Job Title"
            Case "Dept"
                SheetName = "By Department"
                iColumn = 2
                sColumn = "Department"
            Case "DM"
                SheetName = "By Decision Maker"
                iColumn = 4
                sColumn = "Decision Maker"
        End Select
    End If
        
    Dim wk As Worksheet
    Set wk = ActiveWorkbook.Sheets.Add(after:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)) 'Add a worksheet for the analysis using selected SheetName
    wk.Name = SheetName
    Sheets(SheetName).Activate  'Activate new worksheet
    
    '/Minor: START ->
    'Declare and init array containing major demographic category labels for worksheet
    If Minor = 1 Then
        Dim rMinDemos(1 To 6) As String
            rMinDemos(1) = "Asian"
            rMinDemos(2) = "American Indian or Alaskan Native"
            rMinDemos(3) = "Black or African American"
            rMinDemos(4) = "Hispanic"
            rMinDemos(5) = "Native Hawaiian or Other Pacific Islander"
            rMinDemos(6) = "Two or More Races"
    Else
        Dim rMajDemos(1 To 5) As String
            rMajDemos(1) = "Age Over 40"
            rMajDemos(2) = "Female"
            rMajDemos(3) = "Disabled"
            rMajDemos(4) = "Minority"
            rMajDemos(5) = "Veteran"
    End If
    'Minor: <- END\

    Dim iNumValues As Integer  'Number of array elements
    If sAnalysisType = "All" Then
        iNumValues = 1 'Loop only once for "All" - analysis by entire population
        ReDim rString(1 To iNumValues)
            rString(1) = "*"
    Else
        iNumValues = UBound(rUniqueAnalysisValues(sAnalysisType))
        ReDim rString(1 To iNumValues)
            rString = rUniqueAnalysisValues(sAnalysisType)
    End If
      
    Dim f As Integer  'Loop Counter of array index
    Dim fRow As Integer  'Row for data placement
    '/Minor: START ->
    If Minor = 0 Then
        fRow = 2  'First data row (excluding header)
        For f = 1 To iNumValues 'Loop through placement of Worksheet Shell values by the amount of uniquevalues in
            If sAnalysisType <> "All" Then  'Case "All", Analysis Value Column not needed, first Column contains Major Demo Category Labels
                Sheets(SheetName).Cells(fRow, 1).Value = rString(f)  '1 of 5 placements of rAnalysisValue(i)
                Sheets(SheetName).Cells(fRow + 1, 1).Value = rString(f)
                Sheets(SheetName).Cells(fRow + 2, 1).Value = rString(f)
                Sheets(SheetName).Cells(fRow + 3, 1).Value = rString(f)
                Sheets(SheetName).Cells(fRow + 4, 1).Value = rString(f)
            End If
            Sheets(SheetName).Cells(fRow, iColumn).Value = rMajDemos(1)  '1 of 5 placements of each value in rMajDemos
            Sheets(SheetName).Cells(fRow + 1, iColumn).Value = rMajDemos(2)
            Sheets(SheetName).Cells(fRow + 2, iColumn).Value = rMajDemos(3)
            Sheets(SheetName).Cells(fRow + 3, iColumn).Value = rMajDemos(4)
            Sheets(SheetName).Cells(fRow + 4, iColumn).Value = rMajDemos(5)
            fRow = fRow + 6  'Increase starting row count by 6 (5 rows of data + blank row for readability)
        Next
    'Minor: <- END\
    
    Else
        fRow = 2  'First data fRow4 (excluding header)
        For f = 1 To iNumValues 'Loop through placement of Worksheet Shell values by the amount of uniquevalues in
            If sAnalysisTypeMinor <> "All" Then  'Case "All", Analysis Value Column not needed, first Column contains Major Demo Category Labels
                Sheets(SheetName).Cells(fRow, 1).Value = rString(f)  '1 of 5 placements of rAnalysisValue(i)
                Sheets(SheetName).Cells(fRow + 1, 1).Value = rString(f)
                Sheets(SheetName).Cells(fRow + 2, 1).Value = rString(f)
                Sheets(SheetName).Cells(fRow + 3, 1).Value = rString(f)
                Sheets(SheetName).Cells(fRow + 4, 1).Value = rString(f)
                Sheets(SheetName).Cells(fRow + 5, 1).Value = rString(f)
            End If
            Sheets(SheetName).Cells(fRow, iColumn).Value = rMinDemos(1)  '1 of 5 placements of each value in rMajDemos
            Sheets(SheetName).Cells(fRow + 1, iColumn).Value = rMinDemos(2)
            Sheets(SheetName).Cells(fRow + 2, iColumn).Value = rMinDemos(3)
            Sheets(SheetName).Cells(fRow + 3, iColumn).Value = rMinDemos(4)
            Sheets(SheetName).Cells(fRow + 4, iColumn).Value = rMinDemos(5)
            Sheets(SheetName).Cells(fRow + 5, iColumn).Value = rMinDemos(6)
            fRow = fRow + 7  'Increase starting fRow4 count by 6 (5 rows of data + blank fRow4 for readability)
        Next
    End If
    
    'Assign Headers
    'Row.Cell order from iColumn to end of worksheet:
    '>Category, Category Pool, Non-Category Pool, Category Proposed RIF, Non-Category Proposed RIF,
    '>DIF, IRA, SD, IRA < 0.8 and DIF >= 1, DIF >= 1 and SD >= 2, Adverse Impact
    With Sheets(SheetName)
        .Cells(1, 1).Value = sColumn
        .Cells(1, iColumn).Value = "Category"
        .Cells(1, iColumn + 1).Value = "Category Pool"
        .Cells(1, iColumn + 2).Value = "NonCategory Pool"
        .Cells(1, iColumn + 3).Value = "Category Proposed RIF"
        .Cells(1, iColumn + 4).Value = "NonCategory Proposed RIF"
        .Cells(1, iColumn + 5).Value = "DIF"
        .Cells(1, iColumn + 6).Value = "IRA"
        .Cells(1, iColumn + 7).Value = "SD"
        .Cells(1, iColumn + 8).Value = "IRA < 0.8 and DIF >= 1"
        .Cells(1, iColumn + 9).Value = "SD >= 2.0 and DIF >= 1"
        .Cells(1, iColumn + 10).Value = "Adverse Impact"
    End With
   
End Function

'Preformat is called immediately after Variable Declarations in Sub Dia
'Unless <1> and/or <2> are uncommented below, that call will do nothing
Public Sub Preformat()
    Dim nPre As Long
    Dim nPreLast As Long
'///<NOTE:/-----------------------------------------------------------------------------------------------------------------------'
'<1>Field "Group", Column 13, bool values are True if the Employee is a Casual Employee (SAP output value = "Csl" for Field EEG). '
'Casual means that the Employee is temporary or contact or otherwise not qualified to count as a "regular" Employee to be counted '
'within the analysis population. DIA App removes all Group = True records prior to loading Data Table Dset.RifDt into Sheet1.     '
'Code <1> can ben uncommented to do this if being used due to Application FAIL.                                                   '
'<1/                                                                                                                              '
'    nPreLast = nLastRow("Sheet1")                                                                                                '
'    For nPre = nPreLast To 2 Step -1                                                                                             '
'        If Cells(nPre, 13).Value = True Then                                                                                     '
'            Cells(nPre, 13).EntireRow.Delete                                                                                     '
'        End If                                                                                                                   '
'    Next                                                                                                                         '
'/1>                                                                                                                              '
'<2>Fields "Title" and "Dept", Columns 11 & 12, allows for additional analysis by groupings. DIA App removes these Analysis Type  '
'groupings per Default Settings (Settings can be modified by user). Unless otherwise specified by the person providing the data,  '
'values in these columns should be removed to simulate those settings.                                                            '
'Code <2> can be uncommented to do this if being used due to Application FAIL.                                                    '
'<2>                                                                                                                              '
'    nPreLast = nLastRow("Sheet1")                                                                                                '
'    Range(Cells(2, 11), Cells(nPreLast, 11)).Delete                                                                              '
'    Range(Cells(2, 12), Cells(nPreLast, 12)).Delete                                                                              '
'----------------------------------------------------------------------------------------------------------------------------/>///'
End Sub

''''''''''''''''''''''''''''''''''''''
''Public Sub: Report format and save''
''''''''''''''''''''''''''''''''''''''
Public Sub ReportDia()
' Format All Sheets in Workbook
' Request user input for Workbook name
' Save completed Workbook as Analysis Output
    
    UpdateProgress (1)  'the rest will only take a few seconds
    
    'Copy Analysis Data into a new Worksheet named Analysis Data
    'Pull Decision Maker Fname & Lname into analysis Worksheet(s), if needed
    'Delete Sheet1
    Dim dataVerificationWkSht As Worksheet
        Set dataVerificationWkSht = Sheets.Add(after:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
        dataVerificationWkSht.Name = "Analysis Data"
    
    Dim stupidColumnsWontAutoFit As Integer
    Dim runningOutofLetters As Boolean
    
    Sheets("Sheet1").Select
        Cells.Select
        Selection.Copy
    Sheets("Analysis Data").Select
        Cells.Select
        ActiveSheet.Paste
    
    Dim sSheetName As String
    'Vlookup DM Fname and Lname if needed, otherwise remove empty columns
    If Not IsEmpty(Sheets("Sheet1").Cells(2, 14)) And Not IsEmpty(Sheets("Sheet1").Cells(2, 15)) Then
        'DM F & Lname Headers if needed
        Sheets("By Decision Maker").Cells(1, 2).Value = "DM First Name"
        Sheets("By Decision Maker (MinSub)").Cells(1, 2).Value = "DM First Name"
        Sheets("By Decision Maker").Cells(1, 3).Value = "DM Last Name"
        Sheets("By Decision Maker (MinSub)").Cells(1, 3).Value = "DM Last Name"

        'Analysis Data (previously referred to as Data Verification) - remove empty columns
        Dim lastColAnalysisData As Integer
        lastColAnalysisData = iLastColumn("Analysis Data")
        Dim cleanupLoop As Integer
        For cleanupLoop = lastColAnalysisData To 1 Step -1
            If IsEmpty(Sheets("Analysis Data").Cells(2, cleanupLoop)) Then
                Sheets("Analysis Data").Columns(cleanupLoop).EntireColumn.Delete
            End If
        Next
        
        For Min = 0 To Minor  'reappropriating Min for loop, Minor will be 1 if Minority process ran, 0 if not
            If Min = 0 Then
                sSheetName = "By Decision Maker"
            Else
                sSheetName = "By Decision Maker (MinSub)"
            End If
            Dim formatLoopRowsCount As Integer
                formatLoopRowsCount = nLastRow(sSheetName)
            Sheets(sSheetName).Select
            Dim formatRow As Integer  'Loop Counter
            For formatRow = 2 To formatLoopRowsCount
                Cells(formatRow, 2).Activate
                    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],Sheet1!C[12]:C[13],2,FALSE)"
                Cells(formatRow, 3).Activate
                    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],Sheet1!C[11]:c[13],3,FALSE)"
            Next
            
            'Get rid of formulas and #N/A (vlookup for empty values in Column 1)
            Cells.Select
                Selection.Copy
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                        :=False, Transpose:=False
                Selection.Replace What:="#N/A", Replacement:="", LookAt:=xlPart, _
                    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                        ReplaceFormat:=False

            If Min = 0 Then
                sSheetName = "AI Decision Maker"
            Else
                sSheetName = "AI Decision Maker (MinSub)"
            End If
            
            If SheetExists(sSheetName) Then
                formatLoopRowsCount = nLastRow(sSheetName)
                Sheets(sSheetName).Select
                For formatRow = 2 To formatLoopRowsCount
                    Cells(formatRow, 2).Activate
                        ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],Sheet1!C[12]:C[13],2,FALSE)"
                    Cells(formatRow, 3).Activate
                        ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],Sheet1!C[11]:c[13],3,FALSE)"
                Next
                
                'Get rid of formulas and #N/A (vlookup for empty values in Column 1)
                Cells.Select
                    Selection.Copy
                        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False
                    Selection.Replace What:="#N/A", Replacement:="", LookAt:=xlPart, _
                        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                            ReplaceFormat:=False
            End If
        Next
    End If
    
    
'///////////////////////////////////////////////////////PATCH 3.6.2017///////////////////////////////////////////////////////
'//////////////////////Removal of DM Name Columns if first row of DMFname column is not populated////////////////////////////
    If SheetExists("By Decision Maker") = True Then
        If IsEmpty(Sheets("By Decision Maker").Cells(1, 2)) Then
            Sheets("By Decision Maker").Columns(3).EntireColumn.Delete
            Sheets("By Decision Maker").Columns(2).EntireColumn.Delete
        End If
    End If
    
    If SheetExists("By Decision Maker (MinSub)") = True Then
        If IsEmpty(Sheets("By Decision Maker (MinSub)").Cells(1, 2)) Then
            Sheets("By Decision Maker (MinSub)").Columns(3).EntireColumn.Delete
            Sheets("By Decision Maker (MinSub)").Columns(2).EntireColumn.Delete
        End If
    End If
        
    If SheetExists("AI Decision Maker") = True Then
        If IsEmpty(Sheets("AI Decision Maker").Cells(2, 2)) Then
            Sheets("AI Decision Maker").Columns(3).EntireColumn.Delete
            Sheets("AI Decision Maker").Columns(2).EntireColumn.Delete
        End If
    End If
        
    If SheetExists("AI Decision Maker (MinSub)") = True Then
        If IsEmpty(Sheets("AI Decision Maker (MinSub)").Cells(2, 2)) Then
            Sheets("AI Decision Maker (MinSub)").Columns(3).EntireColumn.Delete
            Sheets("AI Decision Maker (MinSub)").Columns(2).EntireColumn.Delete
        End If
    End If
'//////////////////////Removal of DM Name Columns if first row of DMFname column is not populated////////////////////////////
'///////////////////////////////////////////////////////PATCH 3.6.2017///////////////////////////////////////////////////////
      
    'Loop through all Sheets, autofit columns
    'Center the header
    'Change header row color to light grey
    Dim numShts As Integer  'Number of Sheets in workbook
        numShts = Sheets.Count
    Dim shtIndex As Integer  'For Loop
    Dim shtName As String
    For shtIndex = 1 To numShts
    'AutoFit Columns to size
        Sheets(shtIndex).Select
        shtName = ActiveSheet.Name  'Assign ActiveSheet name to string for Function nLastRowColumn call
        
        Rows(1).Select
            Selection.RowHeight = 60
            With Selection
                .HorizontalAlignment = xlCenter  'Center Text
                .VerticalAlignment = xlBottom
                .WrapText = True
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            With Selection
                .VerticalAlignment = xlCenter
                .WrapText = True
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            
            Selection.Font.Bold = True
        
        Dim lastColumn As Integer
            lastColumn = iLastColumn(shtName)
        Range(Cells(1, 1), Cells(1, lastColumn)).Select  'All Data Column Header Cells
            With Selection.Interior
                .Pattern = xlSolid  'Change Color
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = -0.149998474074526
                .PatternTintAndShade = 0
            End With
            Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            With Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
        
        Cells.Select
            Cells.EntireColumn.AutoFit
            
        runningOutofLetters = False
        If Cells(1, 2).Value = "Category Pool" Then
            stupidColumnsWontAutoFit = 2
            runningOutofLetters = True
        ElseIf Cells(1, 3).Value = "Category Pool" Then
            stupidColumnsWontAutoFit = 3
            runningOutofLetters = True
        ElseIf Cells(1, 5).Value = "Category Pool" Then
            stupidColumnsWontAutoFit = 5
            runningOutofLetters = True
        End If

        If runningOutofLetters = True Then
            Columns(stupidColumnsWontAutoFit).Select
                Selection.ColumnWidth = 15
            Columns(stupidColumnsWontAutoFit + 1).Select
                Selection.ColumnWidth = 15
            Columns(stupidColumnsWontAutoFit + 2).Select
                Selection.ColumnWidth = 15
            Columns(stupidColumnsWontAutoFit + 3).Select
                Selection.ColumnWidth = 15
            Columns(stupidColumnsWontAutoFit + 7).Select
                Selection.ColumnWidth = 12
            Columns(stupidColumnsWontAutoFit + 8).Select
                Selection.ColumnWidth = 12
        End If
            
        Cells(1, 1).Select
            
    Next
    
    'Delete DataSource
    Sheets("Sheet1").Delete
    
    'Select the first Worksheet
    'Sheets are in order of creation
    Sheets(1).Select
        Cells(1, 1).Select
        
    'Save Workbook
    MsgBox prompt:="Congratulations, Analysis Report is complete." & vbNewLine _
        & "Please choose a name and location to Save the Report.", Title:="Success!"
    Dim SaveAs As String
        SaveAs = Application.GetSaveAsFilename(FileFilter:="Excel Workbook (*.xlsx), *.xlsx")
    ActiveWorkbook.SaveAs fileName:=SaveAs, FileFormat:=xlWorkbookDefault
    ProgressIndicator.Hide
    Application.Visible = True

End Sub

''''''''''''''''''''''''''''''''''''''''
''Public Sub Progress Indicator Update''
''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateProgress(pct)
' Progress Indicator Update
    
    Application.ScreenUpdating = True
    With ProgressIndicator
        .FrameProgress.Caption = Format(pct, "0%")
        .LabelProgress.Width = pct * (.FrameProgress _
            .Width - 10)
    End With
    DoEvents
    Application.ScreenUpdating = False

End Sub

''''''''''''''''''''''''''''''''
''Public general use Functions''
''''''''''''''''''''''''''''''''
Public Function nLastRow(ByVal SheetName As String) As Long
' Returns the row number of the last row in the worksheet containing data
    Dim rng As Range
    Set rng = Sheets(SheetName).Cells
    nLastRow = rng.Find(What:="*", after:=rng.Cells(1), LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False).row
End Function

Public Function nLastRowofColumn(ByVal SheetName As String, ByVal TempColumn) As Long
' Returns the row number of the last row containing data in the column of the worksheet indicated by params
    Dim rng As Range
    Set rng = Sheets(SheetName).Columns(22)
    nLastRowofColumn = rng.Find(What:="*", after:=rng.Cells(1), LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False).row
End Function

Public Function iLastColumn(ByVal SheetName As String) As Integer
' Returns the column number of the last column in the param row
    Dim rng As Range
    Set rng = Sheets(SheetName).Cells
    iLastColumn = rng.Find(What:="*", after:=rng.Cells(1), LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).Column
End Function

Function SheetExists(shName As String, Optional wb As Workbook) As Boolean
' Returns bool indicating whether Sheets(param shName) exists
    Dim sh As Worksheet
    If wb Is Nothing Then Set wb = ThisWorkbook
On Error Resume Next
    Set sh = wb.Sheets(shName)
On Error GoTo 0
    SheetExists = Not sh Is Nothing
End Function


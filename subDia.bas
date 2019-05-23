Attribute VB_Name = "subDia"
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///                                      Disparate Impact Analysis Report Generator                                         ///
'///                                        Executed by BtnRun_Click from DIA App                                            ///
'///                                                  < v1.0 02.28.2017 >                                                    ///
'///  Purpose: Conduct statistical analysis of Reduction in Force (RIF) selections to determine if any areas of statistical  ///
'///    Adverse Impact exists. Produce an Excel report of statistical analysis results and corresponding list of potentially ///
'///    adversely impacted employees.                                                                                        ///
'///                                                                                                                         ///
'///  Intended Use: VB code in this Workbook is compiled into the Application Assembly as a binary injection. This allows VBA///
'///    features to be used while bypassing Excel's Macro Security restrictions. This will also ensure that both analystical ///
'///    tools always provide the same results.  If DIA App Data Validation process fails this subroutine and corresponding   ///
'///    functions can be used as a failsafe/backup for analysis report creation, provided data for the process is manually   ///
'///    compiled and placed in Sheet1.                                                                                       ///
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///<NOTES/                                                                                                                  ///
'<The DIA Application provided Data Table in Sheet1 will only contain values in the Analysis Type Columns (EE ID, Title, Dept,--
' Decision Maker) if the user has elected to conduct the analysis by those types/fields. Otherwise, all row of those columns----
' will contain a non-empty/null value. Analysis of overall population by EE ID is static.  For the other Analysis Type fields,--
' values can be changed during DIA App data import. Field headers can be changed in the output as needed.>----------------------
' <<As the DIA App data validation process only checks for empty/null values and whether the data type can be converted to a----
'   string, Types can be modified for use with other lateral groupings of employees (e.g. by JG, EEO, Location, etc). Vertical--
'   groupings are also possible but can only be conducted for up to three layers/steps as a time. Such use may be appropriate if
'   the decision making process was distinct yet layered, with overlapping employee populations and/or points in time.>>--------
'-------------------------------------------------------------------------------------------------------------------------------
'<Analysis output consists of a Worksheet containing source data, and a worksheet containing analysis results for each Analysis-
' Type. If any indications of Adverse Impact exist for a Type, a corresponding Worksheet is added to display a listing of all---
' employees potentially impacted. For the overall population analysis, this listing would include all reduction-selected--------
' employees in the Demographic Category showing Adverse Impact. For all others, this listing is further truncated by the value--
' of the associated Analysis Type.>---------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------------------------
'<If the need arises to conduct Major or Minor analysis only, the outer most loop range can be changed from (0 to 1) to (0 to 0)
' or (1 to 1) respectively.>                                                                                          /NOTES>///
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Sub Dia()
'///////////////////////////////////////////////////////////////////////////'
'//     Disparate Impact Analysis for Major Demographic Categories        //'
'// <Major Categories: Age Over 40, Female, Disabled, Minority, Veteran>  //'
'// <Minority SubCategories: Hispanice, Asian, Black or African American, //'
'//  American Indian or Alaskan Native, Native Hawaiian or Other Pacific, //'
'//  Islander, Two or More Races.>                                        //'
'///////////////////////////////////////////////////////////////////////////'
' Note: Code was initially written to separate Major and Minor. Updates   //'
' were made to the Major Sub to accomdate Minor. Comments written for the //'
' first iteration may have been missed, but they still apply with minimual//'
' changes to verbage. Comments have been addded for varaibles triggered by//'
' Minor (2nd Loop of the outer most For Loop). Those comments start with  //'
' 'Minor:'.                                                               //'
'/************************************************************************//'
' Analysis by Overall Population (all employees considered for RIF).      //'
'//***********************************************************************//'
' Additional Analysis by User Selection from DIA App for: Decision Maker  //'
' Job Title, and/or Department.                                           //'
'///**********************************************************************//'
' Decision Maker (DM) Analysis includes 3 value pairings for Analysis     //'
' Value: ID, First Name, Last Name. Fname & Lname are optional. Columns   //'
' housing the data are removed during report format if empty.             //'
'////*********************************************************************//'
' Each Analysis by grouping (represented in code by sAnalysisType)        //'
' produces a separate Worksheet of Analysis Results.                      //'
'/////********************************************************************//'
' A second Worksheet for the relevant Analysis Type is created if the     //'
' Analysis Results indicate Adverse Impact in any of the Analysis         //'
' Categories for an Analysis Value. That Worksheet contains a listing of  //'
' all employees selected for RIF, within the grouping of the relevant     //'
' Analysis Value, if they fit the relevant Demographic Category.          //'
'///////////////////////////////////////////////////////////////////////////'

'******************Variables Declared in order of Init ******************************                                                                                              '
'Analysis Types needed for the report: True = needed, False = not needed.           *
Dim rAnalysis(1 To 4) As Boolean                                                    '
'Array of Analysis Types to be conducted.                                           *
Dim r() As String                                                                   '
'Used as identifier for type of analysis.                                           *
Dim sAnalysisType As String                                                         '
'Temp value assignment for Function calls.                                          *
Dim x As Variant                                                                    '
'Inner loop row skip var                                                            *
Dim skip As Integer                                                                 '
'Assign returned array from Public Function rUniqueAnalysisValues call to local var.*
Dim rAnalysisValues As Variant                                                      '
'Call rDif and assign return array to local variable.                               *
Dim dif As Variant                                                                  '
'Call rImpactRatio and assign return array to local variable.                       *
Dim impactRatio As Variant                                                          '
'Call rStandardDeviation and assign return array to local variable.                 *
Dim sD As Variant                                                                   '
'Call Public Function matrixDemoTotals using rAnalysisValues(n) as sAnalysisByValue.*
Dim matrix As Variant                                                               '
'Assign starting column using sAnalysisType.                                        *
Dim iCol As Integer                                                                 '
'Delcare matrix of AnalysisValue and Impacted Demographic pairs.                    *
Dim matrixImpact As Variant                                                         '
'Last row of new analysis worksheet created by Func CreateSheetshell.               *
Dim lastRow As Long                                                                 '
'Worksheet name for Sheets.Add                                                      *
Dim sht As String                                                                   '
'Worksheet object for Sheets.Add.                                                   *
Dim wrk As Worksheet                                                                '
'Category Header Column assignment for Adverse Impact Employee Listing.             *
Dim c As Integer                                                                    '
'Local variable assignment for matrixImpactedEmployees func call.                   *
Dim impactedEmployees As Variant                                                    '
'Number of Impacted Employees for the Impacted Area.                                *
Dim matrixItems As Integer                                                          '
'************************************************************************************
Dim skipRow As Integer  'For impacted employees loop
Dim index As Integer  'Outer Loop Counter
Dim n As Long  'Inner Loop Counter
Dim i As Integer  'Inner Inner Loop Counter
Dim o As Long  'Inner Inner Inner Loop Counter
Dim p As Long 'Inner Inner Inner Inner Loop Counter

Call Preformat  'Refer to comments for this sub in Module funcDia

    Application.ScreenUpdating = False  'Disable screen updating, allows for faster macro run
    Application.DisplayAlerts = False  'Disable alerts for deleting Sheets containing data and closing the Workbook
    Application.Visible = False  'Hide Excel for user experience purposes

    ProgressIndicator.Show vbModeless  'Show for user experience purposes
    PctProgress = 0
Dim Eq, EqS As Integer 'Allow for independent run of Major or Minor
    Eq = 0
    EqS = 1

'MAJOR***Uncomment to Run Major Demographic Category Analysis Only***MAJOR
'    Eq = 0
'    EqS = 0
'MAJOR***************************************************************MAJOR
    
'MINOR***Uncomment to Run Minority Demographic SubCategories Analysis Only***MINOR
'    Eq = 1
'    EqS = 1
'MINOR***********************************************************************MINOR
    For Minor = Eq To EqS
        'Check for emtpy cell in Row(2) of Analysis Type Columns, emtpy cell indicates emtpy column
        'Assign bools to rAnalysis based on Analysis Type user selections made in DIA App Settings
        rAnalysis(1) = True  'Analysis by All employees considered for RIF selection
        If IsEmpty(Sheets("Sheet1").Cells(2, 14)) Then  'Analysis by Decision Maker
            rAnalysis(2) = False
        Else
            rAnalysis(2) = True
        End If
        
        If IsEmpty(Sheets("Sheet1").Cells(2, 11)) Then  'Analysis by Job Title
            rAnalysis(3) = False
        Else
            rAnalysis(3) = True
        End If
        
        If IsEmpty(Sheets("Sheet1").Cells(2, 12)) Then  'Analysis by Department
            rAnalysis(4) = False
        Else
            rAnalysis(4) = True
        End If
        
        'Assign types of analysis to be conducted to variable size array
        If rAnalysis(1) = True And rAnalysis(2) = False And rAnalysis(3) = False And rAnalysis(4) = False Then
            ReDim r(1 To 1) As String
                r(1) = "All"
        ElseIf rAnalysis(2) = True And rAnalysis(3) = True And rAnalysis(4) = True Then
            ReDim r(1 To 4) As String
                r(1) = "All"
                r(2) = "DM"
                r(3) = "Title"
                r(4) = "Dept"
        ElseIf rAnalysis(2) = True And rAnalysis(3) = True And rAnalysis(4) = False Then
            ReDim r(1 To 3) As String
                r(1) = "All"
                r(2) = "DM"
                r(3) = "Title"
        ElseIf rAnalysis(2) = True And rAnalysis(3) = False And rAnalysis(4) = True Then
            ReDim r(1 To 3) As String
                r(1) = "All"
                r(2) = "DM"
                r(3) = "Dept"
        ElseIf rAnalysis(2) = True And rAnalysis(3) = False And rAnalysis(4) = False Then
            ReDim r(1 To 2) As String
                r(1) = "All"
                r(2) = "DM"
        ElseIf rAnalysis(2) = False And rAnalysis(3) = True And rAnalysis(4) = True Then
            ReDim r(1 To 3) As String
                r(1) = "All"
                r(2) = "Title"
                r(3) = "Dept"
        ElseIf rAnalysis(2) = False And rAnalysis(3) = True And rAnalysis(4) = False Then
            ReDim r(1 To 2) As String
                r(1) = "All"
                r(2) = "Title"
        ElseIf rAnalysis(2) = False And rAnalysis(3) = False And rAnalysis(4) = True Then
            ReDim r(1 To 2) As String
                r(1) = "All"
                r(2) = "Dept"
        End If

        'Each Loop completes an Analysis Worksheet and an AI Employee Listing Worksheet (if needed)
        For index = 1 To UBound(r)  'Loop based on number of analysis types needed
            
            'Progress indicator increased by half the percentage of the loop completed
            PctProgress = PctProgress + (1 / (UBound(r) * (3 - EqS))) 'Percent of the Loops completed
            UpdateProgress (PctProgress)  'Call Public Sub to update ProgressIndicator form
            
            'AnalysisType assignment (local variable used for code readability)
            sAnalysisType = r(index)
            
            'Create analysis worksheet shell and assign unique analysis values to array
            CreateSheetshell (sAnalysisType)
            
            ReDim rAnalysisValues(1 To UBound(rUniqueAnalysisValues(sAnalysisType)))
                rAnalysisValues = rUniqueAnalysisValues(sAnalysisType)
            
            skip = 1  'Inner loop row skip var
            For n = 1 To UBound(rAnalysisValues)  'Loop through the number of items in rAnalysisValues
                  
    '/Minor: Change index from 5 to 6 if for Minority Subcategories\'
                If Minor = 1 Then                                   '
                    Min = 6  '6 Minority SubCategories              '
                Else                                                '
                    Min = 5  '5 Major Categories                    '
                End If                                              '
    '\Minor: Change index from 5 to 6 if for Minority Subcategories/'
    
                x = rAnalysisValues(n)
                ReDim matrix(1 To Min, 1 To 4)
                    matrix = matrixDemoTotals(sAnalysisType, x)  'Assign return matrix to local variable
                   
                If sAnalysisType = "All" Then
                    iCol = 2
                ElseIf sAnalysisType = "DM" Then
                    iCol = 5
                ElseIf sAnalysisType = "Title" Then
                    iCol = 3
                ElseIf sAnalysisType = "Dept" Then
                    iCol = 3
                End If
                
                Sheets(SheetName).Select
                'Totals (iCol to iCol + 3)
                For i = 1 To Min  'Loop through inner index of the matrix (demo categories)
                    'Row count excludes header
                    Cells((i + skip) + ((n - 1) * Min), iCol) = matrix(i, 1) 'Assign matrix value for Category total
                    Cells((i + skip) + ((n - 1) * Min), iCol + 1) = matrix(i, 2) 'NonCategory
                    Cells((i + skip) + ((n - 1) * Min), iCol + 2) = matrix(i, 3) 'CategoryRIF
                    Cells((i + skip) + ((n - 1) * Min), iCol + 3) = matrix(i, 4) 'NonCategoryRIF
                Next
    
                'Dif result (iCol + 4)
                ReDim dif(1 To Min)
                dif = rDif(matrix)
                
                For i = 1 To Min  'Loop through array (Difference result for each demo category)
                    Cells((i + skip) + ((n - 1) * Min), iCol + 4).Value = dif(i)  'Assign Dif value
                Next
        
                'Impact Ratio result (iCol + 5)
                ReDim impactRatio(1 To Min)
                impactRatio = rImpactRatio(matrix)
                For i = 1 To Min  'Loop through array (Impact Ratio result for each demo category)
                    Cells((i + skip) + ((n - 1) * Min), iCol + 5).Value = impactRatio(i)  'Assign Impact Ratio value
                Next
                Columns(iCol + 5).Select
                    Selection.NumberFormat = "0.00"  'Show only upto 2 decimal places
                    
                'Standard Deviation (SD) result (iCol + 6)
                ReDim sD(1 To Min)
                sD = rStandardDeviation(matrix)
                For i = 1 To Min  'Loop through array (SD result for each demo category)
                    Cells((i + skip) + ((n - 1) * Min), iCol + 6).Value = sD(i)  'Assign SD value
                Next
                Columns(iCol + 6).Select
                    Selection.NumberFormat = "0.00"  'Show only upto 2 decimal places
        
                For i = 1 To Min
                    If Cells((i + skip) + ((n - 1) * Min), iCol + 4).Value >= 1 And Cells((i + skip) + ((n - 1) * Min), iCol + 5).Value < 0.8 Then
                        Cells((i + skip) + ((n - 1) * Min), iCol + 7).Value = "X"
                    End If
                    If Cells((i + skip) + ((n - 1) * Min), iCol + 4).Value >= 1 And Cells((i + skip) + ((n - 1) * Min), iCol + 6).Value >= 2 Then
                        Cells((i + skip) + ((n - 1) * Min), iCol + 8).Value = "X"
                    End If
                    If Cells((i + skip) + ((n - 1) * Min), iCol + 7).Value = "X" And Cells((i + skip) + ((n - 1) * Min), iCol + 8).Value = "X" Then
                        Cells((i + skip) + ((n - 1) * Min), iCol + 9).Value = "X"
                    End If
                Next
                
                skip = skip + 1
            Next
        
            '/////////////////////////////////////////////////////////////////////////////////'
            'Following section creates and populates new worksheet listing impacted employees.'
            'Adverse Impact is determined by value "X" in the last of the three Columns.      '
            '/////////////////////////////////////////////////////////////////////////////////'

            'Generate a 2D matrix of analysis value and demo category value for impacted category
            x = 0  'Init inner index for matrix
            lastRow = nLastRow(SheetName)  'SheetName, public & previously assigned by function creating analysis worksheet shell
            'Loop to determine how to size matrixImpact
            For i = 2 To lastRow
                If Cells(i, iCol + 9).Value = "X" Then
                    x = x + 1
                End If
            Next
            
            If x > 0 Then
                ReDim matrixImpact(1 To x, 1 To 2)  'ReDim to correct size
                x = 0  'ReInit for loop
                For i = 2 To lastRow
                    If Cells(i, iCol + 9).Value = "X" Then
                        If sAnalysisType = "All" Then
                            x = x + 1
                            matrixImpact(x, 1) = "*"
                            matrixImpact(x, 2) = Cells(i, iCol - 1).Value
                        Else
                            x = x + 1
                            matrixImpact(x, 1) = Cells(i, 1).Value
                            matrixImpact(x, 2) = Cells(i, iCol - 1).Value
                        End If
                    End If
                Next
    
                Select Case sAnalysisType
                    Case "All"
                        If Minor = 0 Then sht = "AI All Considered" Else sht = "AI All Considered (MinSub)"
                    Case "DM"
                        If Minor = 0 Then sht = "AI Decision Maker" Else sht = "AI Decision Maker (MinSub)"
                    Case "Title"
                        If Minor = 0 Then sht = "AI Job Title" Else sht = "AI Job Title (MinSub)"
                    Case "Dept"
                        If Minor = 0 Then sht = "AI Department" Else sht = "AI Department (MinSub)"
                End Select
                
                'Add a worksheet for the analysis using selected sht name after the last Sheet
                Set wrk = ActiveWorkbook.Sheets.Add(after:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
                    wrk.Name = sht
                
                'Headers
                If sAnalysisType = "All" Then
                    c = 1  'No Analysis Value so Categories are listed in Column 1
                ElseIf sAnalysisType = "DM" Then
                    c = 4  'Analysis Value + 2 data pairs (Fname & Lname) so Cat in Col 4
                Else
                    c = 2  'Analysis Value then Cat
                End If
                'Analysis Value Column Header assignment
                'No value assigned to "All" as no Analysis Values exist
                Select Case sAnalysisType
                    Case "DM"
                        Cells(1, 1).Value = "Decision Maker ID"
                    Case "Title"
                        Cells(1, 1).Value = "Job Title"
                    Case "Dept"
                        Cells(1, 1).Value = "Department"
                End Select
                Cells(1, c).Value = "Category"  'Placements of other columns based on c
                Cells(1, c + 1).Value = "Employee ID"
                Cells(1, c + 2).Value = "First Name"
                Cells(1, c + 3).Value = "Last Name"
                If sAnalysisType = "DM" Then
                    Cells(1, c - 2) = "DM First Name"
                    Cells(1, c - 1) = "DM Last Name"
                End If
                        
                skipRow = 1  'Init takes into account header
                'Data Population
                For o = 1 To x
                    ReDim impactedEmployees(1 To UBound(matrixImpactedEmployees(matrixImpact(o, 1), _
                        sAnalysisType, matrixImpact(o, 2)), 1), 1 To 3)  'ReDim to actual size
                        'matrix of impacted emmployee's ID, Fname, Lname
                        impactedEmployees = matrixImpactedEmployees(matrixImpact(o, 1), sAnalysisType, matrixImpact(o, 2))
    
                    matrixItems = UBound(impactedEmployees, 1)
                    
                    For p = 1 To matrixItems - 1 'Loop through Employees in the returned matrix
                        If sAnalysisType = "DM" Then
                            Sheets(sht).Cells(p + skipRow, c - 3).Value = matrixImpact(o, 1)  'Impact Analysis Value
                        ElseIf sAnalysisType = "Title" Or sAnalysisType = "Dept" Then
                            Sheets(sht).Cells(p + skipRow, c - 1).Value = matrixImpact(o, 1)
                        End If
                        Sheets(sht).Cells(p + skipRow, c).Value = matrixImpact(o, 2)  'Impact Demo category
                        Sheets(sht).Cells(p + skipRow, c + 1).Value = impactedEmployees(p, 1)  'Impacted Employee ID
                        Sheets(sht).Cells(p + skipRow, c + 2).Value = impactedEmployees(p, 2)   'Impacted Employee Fname
                        Sheets(sht).Cells(p + skipRow, c + 3).Value = impactedEmployees(p, 3)  'Impacted Employee Lname
                    Next
                
                    skipRow = skipRow + p  'Account for data already place by previous loops, also adds space between AnalysisValue groupings
                    
                Next
                'If user did not select setting for Employee Name in DIA App, columns will be blank
                If IsEmpty(Cells(2, c + 2)) And IsEmpty(Cells(2, c + 3)) Then
                    Cells(2, c + 3).EntireColumn.Delete
                    Cells(2, c + 2).EntireColumn.Delete
                End If
            End If

        Next
        
    Next

    'Format and Save the Report
    'Application.Quit from Sub
    Call ReportDia
    
End Sub

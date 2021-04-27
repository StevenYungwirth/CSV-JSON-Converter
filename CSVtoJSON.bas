Attribute VB_Name = "CSVtoJSOn"
Option Explicit
Public LastUsedInputFolder As String
Public LastUsedOutputFolder As String

Function GetInputCSV() As String
    'Use the entered file location (if available)
    Dim inputFile As String
    inputFile = Me.Range("Input_Folder").Value2 & Me.Range("Input_File_Name").Value2
    
    'If the entered file name doesn't work or isn't a csv, get the file from a dialog
    If Dir(inputFile) = "" Or InStr(inputFile, ".csv") = 0 Then
        inputFile = GetInputFileFromDialog
    End If
    
    'Quit the procedure if the user didn't select the type of file we need.
    If InStr(inputFile, ".csv") = 0 Then
        MsgBox "No CSV file was provided. Macro has halted."
        End
    End If
    
    'Return the file
    GetInputCSV = inputFile
End Function

Function GetInputFileFromDialog() As String
    'Set the initial folder
    Dim initialFolder As String
    
    If LastUsedInputFolder <> vbNullString Then
        'Start at the last folder selected
        initialFolder = LastUsedInputFolder
    Else
        'Start at the downloads folder
        initialFolder = Environ("USERPROFILE") & "\Downloads\"
        
        'If this folder isn't available, go to the top
        If Dir(initialFolder) = "" Then
            initialFolder = "C:\"
        End If
    End If
    
    'Show a file dialog
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select a CSV File to Convert"
        .AllowMultiSelect = False
        .Filters.Add "CSV Files", "*.csv", 1
        .InitialFileName = initialFolder
        .Show
        
        'Get the selected file
        Dim fileSelected As Variant
        For Each fileSelected In .SelectedItems
            GetInputFileFromDialog = .SelectedItems.Item(1)
        Next fileSelected
        
        'Set the last used folder to be the folder of the selected file
        LastUsedInputFolder = GetFolderFromFilePath(.SelectedItems.Item(1))
    End With
End Function

Function GetFolderFromFilePath(fileName As String) As String
    'Return the parent folder of a full file path
    GetFolderFromFilePath = Left(fileName, InStrRev(fileName, "\"))
End Function

Function GetOutputJSON() As String
    'Get the folder
    Dim folderName As String
    folderName = GetOutputFolder
    
    'Get the file
    Dim fileName As String
    fileName = GetOutputFile
    
    'Return the complete path
    GetOutputJSON = folderName & fileName
End Function

Function GetOutputFolder() As String
    'Use the entered file location (if available)
    Dim outputFile As String
    outputFile = Me.Range("Output_Folder").Value2
    
    'If the entered file name isn't entered or doesn't work, get the file from a dialog
    If outputFile = vbNullString Or Dir(outputFile) = "" Then
        outputFile = GetOutputFolderFromDialog
    End If
End Function

Function GetOutputFolderFromDialog() As String
    'Set the initial folder
    Dim initialFolder As String
    
    If LastUsedOutputFolder <> vbNullString Then
        'Start at the last folder selected
        initialFolder = LastUsedOutputFolder
    Else
        'Start at the downloads folder
        initialFolder = Environ("USERPROFILE") & "\Downloads\"
        
        'If this folder isn't available, go to the top
        If Dir(initialFolder) = "" Then
            initialFolder = "C:\"
        End If
    End If
    
    'Show a file dialog
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select a Folder to Output the JSON Into"
        .AllowMultiSelect = False
        .InitialFileName = initialFolder
        .Show
        
        'Get the selected folder
        Dim folderSelected As Variant
        For Each folderSelected In .SelectedItems
            GetOutputFolderFromDialog = .SelectedItems.Item(1)
        Next folderSelected
        
        'Set the last used folder to be the folder of the selected file
        LastUsedOutputFolder = .SelectedItems.Item(1)
    End With
End Function

Function GetOutputFile() As String
    'Use the entered file name if it's there
    Dim outputFile As String
    outputFile = Me.Range("Output_File_Name").Value2
    
    If outputFile = vbNullString Or HasIllegalCharacters(outputFile) Then
        'There was no entered file name or it has characters not allowed in file names
        Dim inputFile As String
        inputFile = Me.Range("Input_File_Name").Value2
        If inputFile <> vbNullString And Not HasIllegalCharacters(inputFile) Then
            'Use the inputted file name
            outputFile = inputFile & ".json"
        Else
            'Use the default file name
            outputFile = "csvdata.json"
        End If
    Else
        If InStr(outputFile, ".json") = 0 Then
            'Add extension onto file name
            outputFile = outputFile & ".json"
        End If
    End If
    
    'Return the output file name
    GetOutputFile = outputFile
End Function

Function HasIllegalCharacters(strIn As String) As Boolean
    'Set the characters not allowed in file names
    Dim strSpecialChars As String
    strSpecialChars = "~""#%&*:<>?{|}/\[]" & Chr(10) & Chr(13)
    
    'See if the inputted string has any of the characters
    Dim i As Integer
    For i = 1 To Len(strSpecialChars)
        If InStr(strIn, Mid(strSpecialCharacters, i, 1)) > 0 Then
            'The string has an illegal character
            HasIllegalCharacters = True
        End If
    Next i
End Function

Sub ConvertCSVFileToJSONFile(csvFilePath As String, jsonFilePath As String)
    'Determine the next file number available for use
    Dim jsonFileNumber As Integer
    jsonFileNumber = FreeFile
    
    'Open the text file
    Open jsonFilePath For Output As jsonFileNumber
    
    'Write the csv data to the file
    ConvertDataToJSON csvFilePath, jsonFileNumber
    
    'Save and close the file
    Close jsonFileNumber
End Sub

Sub ConvertDataToJSON(csvFile As String, jsonFileNumber As Integer)
    'Open the CSV file for reading
    Dim csvFileNumber As Integer
    csvFileNumber = FreeFile
    Open csvFile For Input As csvFileNumber
    
    Dim fileContent As String
    fileContent = Input(LOF(csvFileNumber), csvFileNumber)
    
    Debug.Print fileContent
    
'    'Start the JSON string
'    Print #jsonFileNumber, "{"
'
'    'Iterate through the CSV's data
'    Dim rw As Integer
'    Dim col As Integer
'    For rw = 2 To UBound(csv.Data, 1)
'        For col = 1 To UBound(csv.Data, 2)
'            Dim colString As String
'            colString = ""
'
'            'Put in the header
'            colString = colString & """" & csv.Data(1, col) & """: "
'
'            'If the data point is a string, put quotes around it
'            If Application.WorksheetFunction.IsText(csv.Data(rw, col)) Then
'                colString = colString & """" & csv.Data(rw, col) & """"
'            Else
'                colString = colString & csv.Data(rw, col)
'            End If
'
'            'Put a comma at the end of the line unless the datapoint is from the last column
'            If col < UBound(csv.Data, 2) Then
'                colString = colString & ","
'            End If
'
'            'Print the data to the JSON file
'            Print #jsonFileNumber, colString
'        Next col
'
'        'End the row's object
'        Dim rowString As String
'        rowString = "}"
'
'        'If this isn't the last row, go to the next object
'        If rw < UBound(csv.Data, 1) Then
'            'Finish this object and go to the next one
'            rowString = rowString & "," & Chr(10) & "{"
'        End If
'
'        'Print the end of the object to the JSON file
'        Print #jsonFileNumber, rowString
'    Next rw
End Sub

Sub OpenCSVToRead()
    Dim filePath As String
    filePath = GetFile
    
    Dim csvFileNumber As Integer
    csvFileNumber = FreeFile
    
    Open filePath For Input As TextFile
    
End Sub

Function GetCSVLine() As String

End Function

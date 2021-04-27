VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub CommandButton1_Click()
    'Get input file path
    Dim csvInputFile As String
    csvInputFile = CSVtoJSOn.GetInputCSV
    
    'Get output file path
    Dim jsonOutputFolder As String
    jsonOutputFile = CSVtoJSOn.GetOutputJSON
    
    'Convert the data
    CSVtoJSOn.ConvertCSVFileToJSONFile csvInputFile, jsonOutputFile
End Sub

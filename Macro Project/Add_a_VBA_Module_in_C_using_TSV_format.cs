using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class AddVbaModuleExample
{
    static void Main()
    {
        // Create a new workbook (empty Excel file)
        Workbook workbook = new Workbook();

        // Get the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Add a procedural (standard) VBA module named "TsvProcessor"
        int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Procedural, "TsvProcessor");

        // Retrieve the newly added module
        VbaModule module = vbaProject.Modules[moduleIndex];

        // Set VBA code that demonstrates how to import a TSV (tab‑separated values) file into the first worksheet
        module.Codes = @"Sub ImportTSV()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(1)
    Dim filePath As String
    filePath = ThisWorkbook.Path & ""\data.tsv""
    ws.QueryTables.Add Connection:=""TEXT;"" & filePath, Destination:=ws.Range(""A1"")
    With ws.QueryTables(1)
        .TextFileParseType = xlDelimited
        .TextFileTabDelimiter = True
        .Refresh BackgroundQuery:=False
    End With
End Sub";

        // Save the workbook as a macro‑enabled file (XLSM)
        workbook.Save("TsvProcessor.xlsm", SaveFormat.Xlsm);
    }
}
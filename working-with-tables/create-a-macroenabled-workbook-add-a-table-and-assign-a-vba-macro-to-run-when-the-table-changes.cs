using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsMacroTableDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook wb = new Workbook();

                // Enable macros for the workbook (required for macro‑enabled files)
                wb.Settings.EnableMacros = true;

                // Access the first worksheet
                Worksheet sheet = wb.Worksheets[0];
                sheet.Name = "DataSheet";

                // Populate sample data (A1:C5)
                sheet.Cells["A1"].PutValue("ID");
                sheet.Cells["B1"].PutValue("Name");
                sheet.Cells["C1"].PutValue("Amount");

                sheet.Cells["A2"].PutValue(1);
                sheet.Cells["B2"].PutValue("Alice");
                sheet.Cells["C2"].PutValue(100);

                sheet.Cells["A3"].PutValue(2);
                sheet.Cells["B3"].PutValue("Bob");
                sheet.Cells["C3"].PutValue(200);

                sheet.Cells["A4"].PutValue(3);
                sheet.Cells["B4"].PutValue("Charlie");
                sheet.Cells["C4"].PutValue(300);

                sheet.Cells["A5"].PutValue(4);
                sheet.Cells["B5"].PutValue("Diana");
                sheet.Cells["C5"].PutValue(400);

                // Add a table (ListObject) covering the data range A1:C5
                int firstRow = 0;      // zero‑based index for row 1
                int firstColumn = 0;   // zero‑based index for column A
                int totalRows = 5;     // rows 1‑5
                int totalColumns = 3;  // columns A‑C

                // Add the ListObject and then set its display name
                sheet.ListObjects.Add(firstRow, firstColumn, totalRows, totalColumns, true);
                var table = sheet.ListObjects[sheet.ListObjects.Count - 1];
                table.DisplayName = "MyTable";   // use DisplayName instead of Name

                // ------------------------------------------------------------
                // Add VBA code that runs when the table changes.
                // The code is placed in the worksheet's code module (document module).
                // ------------------------------------------------------------

                // Create a VBA module associated with the worksheet
                int moduleIndex = wb.VbaProject.Modules.Add(sheet);
                VbaModule sheetModule = wb.VbaProject.Modules[moduleIndex];

                // VBA source code
                string vbaCode = @"
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ExitHandler
    Dim tbl As ListObject
    Set tbl = Me.ListObjects(""MyTable"")
    If Not Intersect(Target, tbl.Range) Is Nothing Then
        Call TableChanged
    End If
ExitHandler:
End Sub

Sub TableChanged()
    MsgBox ""The table 'MyTable' has been modified.""
End Sub
";

                // Assign the VBA code to the module
                sheetModule.Codes = vbaCode;

                // Save the workbook as a macro‑enabled file (lifecycle rule: save)
                string outputPath = "MacroEnabledTable.xlsm";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                wb.Save(outputPath, SaveFormat.Xlsm);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log or display the error
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
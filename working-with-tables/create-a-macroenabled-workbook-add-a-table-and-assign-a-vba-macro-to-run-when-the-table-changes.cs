// Title: Create a macro‑enabled .xlsm workbook with a ListObject and Worksheet_Change VBA event using Aspose.Cells for .NET
// Description: Demonstrates how to generate a macro‑enabled workbook, add an Excel table (ListObject), embed a VBA module with a Worksheet_Change handler that reacts to table edits, and save the file as an .xlsm using Aspose.Cells for C#.
// Keywords: Aspose.Cells C# macro enabled workbook | create .xlsm with VBA | add ListObject Excel table Aspose | Worksheet_Change event VBA Aspose.Cells | embed VBA module C# | save workbook as xlsm | Excel table change macro | Aspose.Cells VBA integration
// Common Searches: How to add a Worksheet_Change VBA event to a ListObject with Aspose.Cells | Create a macro‑enabled .xlsm file and insert a table using C# | Attach a macro to an Excel table change in Aspose.Cells for .NET | Save workbook with VBA code using Aspose.Cells | Add VBA module to Aspose.Cells workbook programmatically
// Developer Intent: Generate a .xlsm workbook, insert a ListObject, embed a VBA Worksheet_Change handler that fires on table modifications, and persist the file with macros enabled.
// Use Cases: Display a message box whenever a row is added or edited in the table. | Log each table change to a separate worksheet or external file via a called subroutine. | Automatically refresh charts or pivot tables when the underlying ListObject data is updated.
// AI Prompts: Write C# code with Aspose.Cells to create a macro‑enabled workbook containing a ListObject named 'SalesData' and a VBA macro that writes change details to a log sheet. | Generate a Worksheet_Change handler that only triggers for a specific ListObject and calls a subroutine 'RefreshPivot' in an Aspose.Cells workbook. | Provide step‑by‑step instructions to add a VBA module to an Aspose.Cells workbook and bind a macro to the ListObject change event.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.Tables;

// Demonstrates how to generate a macro‑enabled workbook, add an Excel table (ListObject), embed a VBA module with a Worksheet_Change handler that reacts to table edits, and save the file as an .xlsm using Aspose.Cells for C#.
class MacroEnabledWorkbookDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and enable macros
            Workbook wb = new Workbook();
            wb.Settings.EnableMacros = true;

            // Get the first worksheet and set a friendly name
            Worksheet ws = wb.Worksheets[0];
            ws.Name = "Data";

            // Populate sample data (including header row)
            ws.Cells["A1"].PutValue("ID");
            ws.Cells["B1"].PutValue("Name");
            ws.Cells["A2"].PutValue(1);
            ws.Cells["B2"].PutValue("Alice");
            ws.Cells["A3"].PutValue(2);
            ws.Cells["B3"].PutValue("Bob");

            // Add a ListObject (Excel table) covering the range A1:B3
            // Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
            int tableIdx = ws.ListObjects.Add(0, 0, 3, 2, true);
            ListObject table = ws.ListObjects[tableIdx];
            table.DisplayName = "MyTable";
            table.ShowTableStyleFirstColumn = true;
            table.ShowTableStyleLastColumn = true;
            table.ShowTableStyleRowStripes = true;

            // Add a VBA module with a Worksheet_Change event handler
            int moduleIdx = wb.VbaProject.Modules.Add(VbaModuleType.Document, "Sheet1");
            VbaModule vbaModule = wb.VbaProject.Modules[moduleIdx];

            // VBA code as a regular string (escaped newlines and quotes)
            string vbaCode = 
                "Private Sub Worksheet_Change(ByVal Target As Range)\r\n" +
                "    On Error GoTo ExitHandler\r\n" +
                "    If Not Intersect(Target, Me.ListObjects(\"\"MyTable\"\").Range) Is Nothing Then\r\n" +
                "        Call MyMacro\r\n" +
                "    End If\r\n" +
                "ExitHandler:\r\n" +
                "End Sub\r\n\r\n" +
                "Sub MyMacro()\r\n" +
                "    MsgBox \"\"Table changed!\"\"\r\n" +
                "End Sub";

            vbaModule.Codes = vbaCode;

            // Save the workbook as a macro‑enabled file
            string outputPath = "MacroEnabledTable.xlsm";
            wb.Save(outputPath, SaveFormat.Xlsm);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

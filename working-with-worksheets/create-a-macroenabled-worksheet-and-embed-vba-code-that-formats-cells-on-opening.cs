// Title: Create a macro‑enabled .xlsm workbook with a Workbook_Open VBA macro using Aspose.Cells for .NET (C#)
// Description: Shows how to generate an .xlsm file with Aspose.Cells, add or locate the ThisWorkbook VBA module, inject a Workbook_Open routine that makes range A1:B2 bold with a green background, and save the workbook as a macro‑enabled file.
// Keywords: Aspose.Cells | C# | VBA module | macro‑enabled workbook | xlsm generation | Workbook_Open event | programmatic Excel formatting | embed VBA code | save as .xlsm
// Common Searches: Aspose.Cells add ThisWorkbook VBA module C# | Create .xlsm file with Aspose.Cells | Set Workbook_Open macro using Aspose.Cells | Programmatically format cells on workbook open | Embed VBA code in Excel file with Aspose
// Developer Intent: Generate an .xlsm workbook and embed a Workbook_Open macro that automatically formats a specific cell range when the file is opened.
// Use Cases: Distribute a template that applies consistent header styling on open. | Automate report formatting without requiring user interaction. | Create a self‑formatting Excel tool that enforces visual standards via VBA.
// AI Prompts: Write C# code with Aspose.Cells to create an .xlsm workbook, add a ThisWorkbook module, and insert a Workbook_Open macro that bolds and colors range A1:B2. | Explain how to check for an existing VBA module and add one if missing when generating a macro‑enabled workbook with Aspose.Cells. | Show how to modify the injected VBA code to change the target range or formatting colors using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsMacroExample
{
    // Shows how to generate an .xlsm file with Aspose.Cells, add or locate the ThisWorkbook VBA module, inject a Workbook_Open routine that makes range A1:B2 bold with a green background, and save the workbook as a macro‑enabled file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (default format is XLSX)
                Workbook workbook = new Workbook();

                // Access the VBA project (creates a project when saved as macro‑enabled)
                VbaProject vbaProject = workbook.VbaProject;

                // Try to get the existing "ThisWorkbook" module; if it does not exist, add it
                VbaModule thisWorkbookModule = null;
                foreach (VbaModule mod in vbaProject.Modules)
                {
                    if (mod.Name.Equals("ThisWorkbook", StringComparison.OrdinalIgnoreCase))
                    {
                        thisWorkbookModule = mod;
                        break;
                    }
                }

                if (thisWorkbookModule == null)
                {
                    int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Document, "ThisWorkbook");
                    thisWorkbookModule = vbaProject.Modules[moduleIndex];
                }

                // Set VBA code for the Workbook_Open event
                thisWorkbookModule.Codes =
                    "Private Sub Workbook_Open()\r\n" +
                    "    With ThisWorkbook.Worksheets(1).Range(\"A1:B2\")\r\n" +
                    "        .Font.Bold = True\r\n" +
                    "        .Interior.Color = &H00FF00 ' Green background\r\n" +
                    "    End With\r\n" +
                    "End Sub";

                // Save the workbook as a macro‑enabled file (.xlsm)
                workbook.Save("MacroEnabledWorkbook.xlsm", SaveFormat.Xlsm);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}

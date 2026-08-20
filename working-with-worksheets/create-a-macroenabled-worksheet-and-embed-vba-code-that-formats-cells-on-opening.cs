// Title: Create a macro‑enabled .xlsm workbook with VBA Workbook_Open formatting using Aspose.Cells C#
// Description: C# example that builds a new Workbook, accesses its VbaProject, ensures a Document module (ThisWorkbook) and a Procedural module (HelperModule) exist, injects a Workbook_Open macro that colors range A1:B2 yellow and makes the font bold, adds a ShowMessage subroutine, and saves the file as a macro‑enabled .xlsm workbook.
// Keywords: Aspose.Cells | C# | VBA | macro‑enabled workbook | xlsm | Workbook_Open event | add VBA module | embed VBA code | format cells on open | VbaProject | Excel automation
// Common Searches: Aspose.Cells add Workbook_Open macro C# | create .xlsm file with VBA using Aspose.Cells | how to embed VBA modules in Excel workbook with Aspose.Cells | save macro‑enabled workbook Aspose.Cells .NET | format cells on workbook open Aspose.Cells example
// Developer Intent: Generate a macro‑enabled Excel file and programmatically embed VBA code that formats cells when the workbook is opened.
// Use Cases: Automatically apply a yellow background and bold font to a specific range each time the file is opened. | Display a confirmation message box on workbook open to verify macro execution. | Organize VBA logic by separating event code (ThisWorkbook) from reusable procedures (HelperModule).
// AI Prompts: Write C# code with Aspose.Cells that adds a Workbook_Open macro to set A1:B2 background to yellow and font to bold. | Show how to create a procedural VBA module named HelperModule containing a public Sub ShowMessage that displays a message box, then save the workbook as .xlsm. | Explain how to check for existing VBA modules in a VbaProject before adding new ones using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

// C# example that builds a new Workbook, accesses its VbaProject, ensures a Document module (ThisWorkbook) and a Procedural module (HelperModule) exist, injects a Workbook_Open macro that colors range A1:B2 yellow and makes the font bold, adds a ShowMessage subroutine, and saves the file as a macro‑enabled .xlsm workbook.
class MacroWorkbookDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook (default format is XLSX)
            Workbook wb = new Workbook();

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = wb.VbaProject;

            // Get existing Document module named "ThisWorkbook" or add a new one
            VbaModule docModule = null;
            foreach (VbaModule mod in vbaProject.Modules)
            {
                if (mod.Type == VbaModuleType.Document && mod.Name.Equals("ThisWorkbook", StringComparison.OrdinalIgnoreCase))
                {
                    docModule = mod;
                    break;
                }
            }
            if (docModule == null)
            {
                int docModuleIndex = vbaProject.Modules.Add(VbaModuleType.Document, "ThisWorkbook");
                docModule = vbaProject.Modules[docModuleIndex];
            }

            // VBA code that runs when the workbook is opened and formats cells A1:B2
            string workbookOpenCode =
                "Private Sub Workbook_Open()\r\n" +
                "    Dim ws As Worksheet\r\n" +
                "    Set ws = ThisWorkbook.Worksheets(1)\r\n" +
                "    ws.Range(\"A1:B2\").Interior.Color = RGB(255, 255, 0) ' Yellow background\r\n" +
                "    ws.Range(\"A1:B2\").Font.Bold = True\r\n" +
                "End Sub";

            docModule.Codes = workbookOpenCode;

            // Get existing Procedural module named "HelperModule" or add a new one
            VbaModule procModule = null;
            foreach (VbaModule mod in vbaProject.Modules)
            {
                if (mod.Type == VbaModuleType.Procedural && mod.Name.Equals("HelperModule", StringComparison.OrdinalIgnoreCase))
                {
                    procModule = mod;
                    break;
                }
            }
            if (procModule == null)
            {
                int procModuleIndex = vbaProject.Modules.Add(VbaModuleType.Procedural, "HelperModule");
                procModule = vbaProject.Modules[procModuleIndex];
            }

            // Auxiliary subroutine
            procModule.Codes = "Public Sub ShowMessage()\r\n    MsgBox \"Workbook opened\"\r\nEnd Sub";

            // Save the workbook as a macro‑enabled file (.xlsm)
            wb.Save("MacroEnabledWorkbook.xlsm", SaveFormat.Xlsm);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}

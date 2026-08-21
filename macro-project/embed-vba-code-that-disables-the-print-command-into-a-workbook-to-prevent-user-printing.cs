// Title: Add VBA to Block Printing in an Excel Workbook with Aspose.Cells (C#)
// Description: Creates a new workbook, injects a Workbook_BeforePrint handler that cancels the print job and shows a warning, optionally protects the VBA project, and saves the file as a macro‑enabled .xlsm workbook.
// Keywords: Aspose.Cells C# VBA injection | Workbook_BeforePrint event | prevent Excel printing | macro‑enabled XLSM generation | protect VBA project programmatically | Excel security with Aspose
// Common Searches: Aspose.Cells add VBA to stop printing | C# embed Workbook_BeforePrint macro | save macro enabled workbook with protected VBA | prevent users from printing Excel file using .NET | how to inject VBA code with Aspose.Cells
// Developer Intent: Insert a VBA routine that blocks printing and save the workbook as a macro‑enabled file.
// Use Cases: Enforce a no‑print policy for confidential spreadsheets. | Distribute Excel templates that automatically warn and block printing. | Generate reports that require VBA protection while disabling the print command.
// AI Prompts: Write C# code using Aspose.Cells to add a Workbook_BeforePrint procedure that cancels printing and displays a message box. | Show how to protect the VBA project with a password while keeping the code accessible in a generated .xlsm file. | Explain how to verify that the ThisWorkbook module contains the expected VBA before saving the workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// Creates a new workbook, injects a Workbook_BeforePrint handler that cancels the print job and shows a warning, optionally protects the VBA project, and saves the file as a macro‑enabled .xlsm workbook.
class DisablePrintVbaDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the existing ThisWorkbook class module (it is added by default)
            VbaModule vbaModule = null;
            foreach (VbaModule module in workbook.VbaProject.Modules)
            {
                if (module.Name.Equals("ThisWorkbook", StringComparison.OrdinalIgnoreCase))
                {
                    vbaModule = module;
                    break;
                }
            }

            // If for some reason it does not exist, add it (fallback)
            if (vbaModule == null)
            {
                int moduleIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Class, "ThisWorkbook");
                vbaModule = workbook.VbaProject.Modules[moduleIndex];
            }

            // VBA code that cancels any print attempt
            string vbaCode = @"
Private Sub Workbook_BeforePrint(Cancel As Boolean)
    Cancel = True
    MsgBox ""Printing is disabled by VBA.""
End Sub
";

            // Assign the VBA code to the module
            vbaModule.Codes = vbaCode;

            // Optionally protect the VBA project (not locked for viewing)
            workbook.VbaProject.Protect(false, "vbaPassword");

            // Define output path
            string outputPath = "DisablePrint.xlsm";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a macro‑enabled file
            workbook.Save(outputPath, SaveFormat.Xlsm);

            Console.WriteLine($"Workbook saved to {outputPath} with VBA that disables printing.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

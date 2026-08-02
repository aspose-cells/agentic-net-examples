// Title: C# – Batch Add a Printing‑Prevention VBA Macro to All .xlsm Files with Aspose.Cells
// Description: A C# console app that scans a folder for macro‑enabled Excel workbooks (*.xlsm), loads each file with Aspose.Cells, creates or updates the "ThisWorkbook" VBA class, inserts a Workbook_BeforePrint routine that cancels printing and shows a message, then saves the workbook back in Xlsm format, overwriting the original file.
// Keywords: Aspose.Cells | C# add VBA macro | batch process xlsm | prevent printing Excel | Workbook_BeforePrint | macro‑enabled workbook | VBA project manipulation | automate VBA insertion | Excel automation .NET | Windows C# Excel macro
// Common Searches: add beforeprint macro to multiple xlsm files using C# | batch insert VBA into Excel workbooks Aspose.Cells | disable printing in Excel programmatically | C# code to update ThisWorkbook module in all workbooks | overwrite macro‑enabled workbook with new VBA code
// Developer Intent: Insert a VBA routine that blocks printing into every macro‑enabled workbook within a specified directory, using Aspose.Cells for .NET.
// Use Cases: Enforce a no‑print policy on confidential Excel templates before distribution. | Automatically embed a printing‑prevention macro into generated reports to protect sensitive data. | Upgrade existing macro‑enabled workbooks with a security macro without manual editing.
// AI Prompts: Generate C# code with Aspose.Cells that adds a Workbook_BeforePrint macro to all .xlsm files in a folder and overwrites each file. | Show how to detect whether the "ThisWorkbook" module already contains the printing‑prevention code before replacing it. | Create a version of the script that logs processed file names and any errors to a CSV file while adding the macro.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// A C# console app that scans a folder for macro‑enabled Excel workbooks (*.xlsm), loads each file with Aspose.Cells, creates or updates the "ThisWorkbook" VBA class, inserts a Workbook_BeforePrint routine that cancels printing and shows a message, then saves the workbook back in Xlsm format, overwriting the original file.
class BatchAddPrintPreventionMacro
{
    static void Main()
    {
        // Folder containing macro‑enabled workbooks (*.xlsm)
        string folderPath = @"C:\Path\To\MacroWorkbooks";

        // Verify that the folder exists
        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine($"Folder not found: {folderPath}");
            return;
        }

        // Get all .xlsm files in the folder
        string[] files;
        try
        {
            files = Directory.GetFiles(folderPath, "*.xlsm");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error retrieving files: {ex.Message}");
            return;
        }

        foreach (string filePath in files)
        {
            // Ensure the file still exists before processing
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found (skipped): {filePath}");
                continue;
            }

            try
            {
                // Load the workbook (macro‑enabled)
                Workbook workbook = new Workbook(filePath);

                // Add or replace a VBA class module named "ThisWorkbook"
                int moduleIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Class, "ThisWorkbook");
                VbaModule vbaModule = workbook.VbaProject.Modules[moduleIndex];
                vbaModule.Codes =
@"Private Sub Workbook_BeforePrint(Cancel As Boolean)
    Cancel = True
    MsgBox ""Printing is disabled.""
End Sub";

                // Save the workbook back as macro‑enabled file (overwrites original)
                workbook.Save(filePath, SaveFormat.Xlsm);
                Console.WriteLine($"Processed: {Path.GetFileName(filePath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("Processing completed. Printing‑prevention macro added to all workbooks.");
    }
}

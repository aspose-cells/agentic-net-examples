// Title: Batch add a Workbook_BeforePrint VBA macro to all .xlsm files in a directory using Aspose.Cells for .NET
// AI Prompts: Write C# code that iterates through every .xlsm file in a given folder, loads each workbook with Aspose.Cells, and injects a Workbook_BeforePrint event macro that cancels printing while preserving existing macros. | Enhance the batch macro insertion script to produce a log that records for each workbook whether the printing‑prevention macro was added, already existed, or was skipped because the file lacked a VBA project, using the Aspose.Cells API. | Create a reusable C# method that receives a workbook path and a VBA code snippet, checks for the ThisWorkbook module, appends the code only if the specified procedure is missing, and saves the file back as an .xlsm workbook.
// Common Searches: how to programmatically add a Workbook_BeforePrint event to multiple xlsm workbooks using Aspose.Cells C# | batch insert VBA macro that disables printing into macro-enabled Excel files with .NET | Aspose.Cells C# example for updating VBA project in existing .xlsm files | C# script to process a folder of macro-enabled Excel workbooks and add printing protection macro | skip Excel files without VBA project when adding macros with Aspose.Cells
// Tags: batch add VBA macro Aspose.Cells .xlsm | Workbook_BeforePrint event injection C# | update VBA project programmatically Aspose.Cells | process macro-enabled workbooks folder .NET | prevent Excel printing via VBA macro Aspose

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;   // Required for VbaProject and VbaModule

// The C# program scans a specified directory for .xlsm workbooks, loads each file with Aspose.Cells while preserving existing macros, verifies that a VBA project and the ThisWorkbook module are present, and appends a Workbook_BeforePrint event macro that cancels printing if it is not already defined. Each workbook is saved back in Xlsm format, and the script logs which files were modified, skipped, or caused errors.
class MacroBatchProcessor
{
    static void Main()
    {
        // Path to the folder containing macro‑enabled workbooks (*.xlsm)
        string folderPath = @"C:\Path\To\Folder";

        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine($"Folder not found: {folderPath}");
            return;
        }

        // VBA code that disables printing by handling the Workbook_BeforePrint event
        const string printingPreventionCode = @"
Private Sub Workbook_BeforePrint(Cancel As Boolean)
    MsgBox ""Printing is disabled by policy.""
    Cancel = True
End Sub
";

        // Process each .xlsm file in the folder
        foreach (string filePath in Directory.GetFiles(folderPath, "*.xlsm"))
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found (skipped): {filePath}");
                    continue;
                }

                // Load the workbook (preserves existing macros)
                Workbook workbook = new Workbook(filePath);

                // Ensure a VBA project exists
                VbaProject vbaProject = workbook.VbaProject;
                if (vbaProject == null)
                {
                    Console.WriteLine($"No VBA project in file (skipped): {filePath}");
                    continue;
                }

                // Find the ThisWorkbook module (contains workbook‑level events)
                VbaModule thisWorkbookModule = null;
                foreach (VbaModule module in vbaProject.Modules)
                {
                    if (module.Name.Equals("ThisWorkbook", StringComparison.OrdinalIgnoreCase))
                    {
                        thisWorkbookModule = module;
                        break;
                    }
                }

                if (thisWorkbookModule == null)
                {
                    Console.WriteLine($"ThisWorkbook module missing (skipped): {filePath}");
                    continue;
                }

                // Add the event code if it is not already present
                if (!thisWorkbookModule.Codes.Contains("Workbook_BeforePrint"))
                {
                    thisWorkbookModule.Codes += printingPreventionCode;
                }

                // Save the workbook back as a macro‑enabled file
                workbook.Save(filePath, SaveFormat.Xlsm);
                Console.WriteLine($"Processed: {Path.GetFileName(filePath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("Processing completed.");
    }
}

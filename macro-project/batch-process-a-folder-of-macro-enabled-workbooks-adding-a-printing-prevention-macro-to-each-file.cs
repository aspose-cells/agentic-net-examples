using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class BatchAddPrintPreventionMacro
{
    static void Main()
    {
        try
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
            string[] files = Directory.GetFiles(folderPath, "*.xlsm", SearchOption.TopDirectoryOnly);

            foreach (string file in files)
            {
                // Ensure the file still exists before processing
                if (!File.Exists(file))
                {
                    Console.WriteLine($"File not found (skipped): {file}");
                    continue;
                }

                try
                {
                    // Load the workbook (macro‑enabled)
                    Workbook workbook = new Workbook(file);

                    // Ensure the workbook has a VBA project; if not, create one by saving as .xlsm and reloading
                    if (workbook.VbaProject == null || workbook.VbaProject.Modules.Count == 0)
                    {
                        workbook.Save(file, SaveFormat.Xlsm);
                        workbook = new Workbook(file);
                    }

                    // Locate the ThisWorkbook module (where workbook events are placed)
                    int thisWorkbookIndex = -1;
                    for (int i = 0; i < workbook.VbaProject.Modules.Count; i++)
                    {
                        if (workbook.VbaProject.Modules[i].Name.Equals("ThisWorkbook", StringComparison.OrdinalIgnoreCase))
                        {
                            thisWorkbookIndex = i;
                            break;
                        }
                    }

                    // If ThisWorkbook module does not exist, add it as a class module
                    if (thisWorkbookIndex == -1)
                    {
                        thisWorkbookIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Class, "ThisWorkbook");
                    }

                    // Insert the Workbook_BeforePrint event to cancel printing
                    VbaModule thisWorkbookModule = workbook.VbaProject.Modules[thisWorkbookIndex];
                    thisWorkbookModule.Codes =
@"Private Sub Workbook_BeforePrint(Cancel As Boolean)
    MsgBox ""Printing is disabled by policy.""
    Cancel = True
End Sub";

                    // Save the modified workbook back to the same file (keep macro format)
                    workbook.Save(file, SaveFormat.Xlsm);
                    Console.WriteLine($"Processed: {Path.GetFileName(file)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{file}': {ex.Message}");
                }
            }

            Console.WriteLine("Processing completed. Printing‑prevention macro added to all workbooks.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
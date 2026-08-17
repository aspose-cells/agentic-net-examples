// Title: Batch add a print‑blocking VBA macro to all .xlsm files with Aspose.Cells C#
// Description: Iterates through every *.xlsm workbook in a specified folder, loads each file with Aspose.Cells, confirms a VBA project exists, injects a Workbook_BeforePrint routine into the ThisWorkbook class to display a warning and cancel printing, then overwrites the file in macro‑enabled format.
// Keywords: Aspose.Cells VBA macro injection | C# batch process xlsm | prevent printing Excel macro | add ThisWorkbook module programmatically | save macro‑enabled workbook .NET | automate VBA code insertion | Excel print disable Aspose
// Common Searches: how to add a VBA macro to multiple xlsm files using Aspose.Cells | C# code to disable printing in Excel workbooks programmatically | batch insert Workbook_BeforePrint event with Aspose.Cells | add ThisWorkbook class module to existing VBA project .NET | prevent Excel printing across a folder of macro‑enabled files
// Developer Intent: Insert a VBA routine that blocks printing into every macro‑enabled workbook in a given directory using Aspose.Cells for .NET.
// Use Cases: Enforce a no‑print policy on a collection of Excel templates before distribution. | Update legacy .xlsm reports with a print‑cancellation macro without opening Excel manually. | Integrate into CI/CD pipelines to guarantee all generated macro‑enabled files contain the print‑prevention code.
// AI Prompts: Write C# code with Aspose.Cells that adds a Workbook_BeforePrint macro to all .xlsm files in a folder, handling missing VBA projects gracefully. | Show how to modify or create the ThisWorkbook class module and inject VBA code using the Aspose.Cells VbaProject API. | Provide best‑practice error handling and logging for batch processing of macro‑enabled workbooks with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace MacroProcessing
{
    // Iterates through every *.xlsm workbook in a specified folder, loads each file with Aspose.Cells, confirms a VBA project exists, injects a Workbook_BeforePrint routine into the ThisWorkbook class to display a warning and cancel printing, then overwrites the file in macro‑enabled format.
    class BatchAddPrintPreventionMacro
    {
        static void Main()
        {
            // Folder containing macro‑enabled workbooks (*.xlsm)
            string folderPath = @"C:\Path\To\MacroWorkbooks";

            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Folder not found: {folderPath}");
                return;
            }

            // Iterate over each .xlsm file in the folder
            foreach (string filePath in Directory.GetFiles(folderPath, "*.xlsm"))
            {
                try
                {
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found: {filePath}");
                        continue;
                    }

                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);

                    // Ensure the workbook has a VBA project
                    VbaProject vbaProject = workbook.VbaProject;
                    if (vbaProject == null)
                    {
                        Console.WriteLine($"No VBA project found in: {Path.GetFileName(filePath)}. Skipping.");
                        continue;
                    }

                    // Add or reuse the class module named "ThisWorkbook"
                    int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Class, "ThisWorkbook");
                    VbaModule module = vbaProject.Modules[moduleIndex];

                    // VBA code that cancels any print operation
                    module.Codes =
                        "Private Sub Workbook_BeforePrint(Cancel As Boolean)\r\n" +
                        "    MsgBox \"Printing is disabled.\"\r\n" +
                        "    Cancel = True\r\n" +
                        "End Sub";

                    // Overwrite the original file with the macro added
                    workbook.Save(filePath, SaveFormat.Xlsm);
                    Console.WriteLine($"Processed: {Path.GetFileName(filePath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }
        }
    }
}

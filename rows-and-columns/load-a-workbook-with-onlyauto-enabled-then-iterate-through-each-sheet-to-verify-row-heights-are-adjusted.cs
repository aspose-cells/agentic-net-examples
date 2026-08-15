// Title: C# – Load Workbook with OnlyAuto, AutoFitRows per Sheet, Verify Row Height (Aspose.Cells)
// Description: Loads an Excel file using LoadOptions with AutoFitterOptions.OnlyAuto, iterates through every worksheet, applies AutoFitRows respecting the OnlyAuto flag, checks each row's IsHeightMatched property, and saves the result.
// Keywords: Aspose.Cells OnlyAuto | AutoFitRows C# | Row.IsHeightMatched | load workbook with AutoFitterOptions | verify auto‑fit rows | preserve custom row height | C# Excel row height check
// Common Searches: Aspose.Cells enable OnlyAuto when loading workbook | C# AutoFitRows only for rows without custom height | how to check Row.IsHeightMatched after AutoFitRows | iterate worksheets and verify row heights Aspose.Cells
// Developer Intent: Load a workbook with OnlyAuto, auto‑fit rows on each sheet, and identify which rows were resized.
// Use Cases: Adjust row heights automatically while keeping manually set heights unchanged. | Create a fallback workbook when the source file is missing, then apply OnlyAuto auto‑fit to its default sheet. | Log Row.IsHeightMatched for every row to confirm auto‑fit results before saving.
// AI Prompts: Write C# code that loads an Excel file with LoadOptions.OnlyAuto, calls AutoFitRows on each worksheet, and prints Row.IsHeightMatched for all rows. | Explain the meaning of Row.IsHeightMatched after using AutoFitRows with the OnlyAuto option in Aspose.Cells. | Provide a step‑by‑step guide to verify that only rows without custom heights are resized when OnlyAuto is enabled.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsOnlyAutoDemo
{
    // Loads an Excel file using LoadOptions with AutoFitterOptions.OnlyAuto, iterates through every worksheet, applies AutoFitRows respecting the OnlyAuto flag, checks each row's IsHeightMatched property, and saves the result.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the source workbook
                string inputPath = "input.xlsx";

                // Ensure the input file exists; create a simple workbook if it does not
                Workbook workbook;
                if (File.Exists(inputPath))
                {
                    // Configure load options with OnlyAuto enabled
                    LoadOptions loadOptions = new LoadOptions
                    {
                        AutoFitterOptions = new AutoFitterOptions
                        {
                            OnlyAuto = true // Fit only rows that do not have custom height
                        }
                    };

                    // Load the workbook using the configured options
                    workbook = new Workbook(inputPath, loadOptions);
                }
                else
                {
                    // Create a new workbook with a sample sheet when the input file is missing
                    workbook = new Workbook();
                    Worksheet sheet = workbook.Worksheets[0];
                    sheet.Name = "SampleSheet";
                    sheet.Cells["A1"].PutValue("Sample data");
                }

                // Prepare AutoFitterOptions that will be used for explicit AutoFitRows calls
                AutoFitterOptions fitOptions = new AutoFitterOptions
                {
                    OnlyAuto = true
                };

                // Iterate through each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Auto‑fit rows in the current sheet respecting OnlyAuto option
                    sheet.AutoFitRows(fitOptions);

                    // Verify row heights: rows whose height matches default font are considered auto‑fitted
                    int maxRow = sheet.Cells.MaxDataRow; // last row with data
                    for (int i = 0; i <= maxRow; i++)
                    {
                        Row row = sheet.Cells.Rows[i];
                        // Output verification result
                        Console.WriteLine($"Sheet \"{sheet.Name}\", Row {i + 1}: IsHeightMatched = {row.IsHeightMatched}");
                    }
                }

                // Save the modified workbook
                string outputPath = "output.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

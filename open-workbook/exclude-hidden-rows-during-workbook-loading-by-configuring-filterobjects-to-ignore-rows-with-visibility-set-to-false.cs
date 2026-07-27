// Title: C# – Load an Excel workbook without hidden rows using Aspose.Cells LoadFilter
// Description: Demonstrates how to create a custom LoadFilter (VisibleRowsLoadFilter) that loads all data, then iterates each worksheet, deletes rows where Cells.IsRowHidden is true, and saves a clean workbook. The example also shows how to generate a sample file with a hidden row if none exists.
// Keywords: Aspose.Cells LoadFilter hidden rows | C# exclude hidden rows Excel | LoadOptions filter hidden rows | remove hidden rows Aspose.Cells | skip hidden rows during load | Aspose.Cells workbook cleaning
// Common Searches: how to ignore hidden rows when loading Excel with Aspose.Cells .NET | Aspose.Cells custom LoadFilter example | remove hidden rows after loading workbook C# | load workbook without hidden rows Aspose | skip hidden rows using LoadOptions
// Developer Intent: Load an Excel file and automatically discard any rows that are hidden, producing a workbook that contains only visible data.
// Use Cases: Import a template that contains hidden helper rows and export a clean version for downstream processing. | Validate user‑uploaded spreadsheets while ensuring hidden rows do not affect calculations or reports. | Generate data extracts where hidden rows must be omitted to meet compliance or presentation requirements.
// AI Prompts: Show a C# Aspose.Cells snippet that uses LoadFilter to prevent hidden rows from being loaded at all. | Explain how to extend VisibleRowsLoadFilter to also skip hidden columns during workbook loading. | Suggest an alternative approach that uses LoadOptions.FilterObjects to exclude hidden rows before they are read into memory.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Custom LoadFilter that loads all data (you can adjust options if needed)
    // Demonstrates how to create a custom LoadFilter (VisibleRowsLoadFilter) that loads all data, then iterates each worksheet, deletes rows where Cells.IsRowHidden is true, and saves a clean workbook. The example also shows how to generate a sample file with a hidden row if none exists.
    public class VisibleRowsLoadFilter : LoadFilter
    {
        public override void StartSheet(Worksheet sheet)
        {
            // Load everything for the sheet; hidden rows will be filtered out later
            LoadDataFilterOptions = LoadDataFilterOptions.All;
        }
    }

    public class ExcludeHiddenRowsDemo
    {
        public static void Run()
        {
            // Path to the source workbook (contains hidden rows)
            string sourcePath = "SourceWithHiddenRows.xlsx";

            // Ensure the source file exists; create a sample workbook if it does not
            if (!File.Exists(sourcePath))
            {
                CreateSampleWorkbook(sourcePath);
            }

            try
            {
                // Configure load options with the custom filter
                LoadOptions loadOptions = new LoadOptions
                {
                    LoadFilter = new VisibleRowsLoadFilter()
                };

                // Load the workbook using the configured options
                Workbook workbook = new Workbook(sourcePath, loadOptions);

                // Iterate through each worksheet and remove rows that are hidden
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Process rows in reverse order to avoid index shifting when deleting
                    for (int row = sheet.Cells.MaxDataRow; row >= sheet.Cells.MinDataRow; row--)
                    {
                        if (sheet.Cells.IsRowHidden(row))
                        {
                            // Delete the hidden row so it is excluded from the loaded workbook
                            sheet.Cells.DeleteRow(row);
                        }
                    }
                }

                // Save the resulting workbook; hidden rows are no longer present
                string outputPath = "ResultWithoutHiddenRows.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
            }
        }

        // Helper method to create a sample workbook with a hidden row
        private static void CreateSampleWorkbook(string path)
        {
            try
            {
                Workbook wb = new Workbook();
                Worksheet ws = wb.Worksheets[0];
                ws.Cells["A1"].PutValue("Row 1");
                ws.Cells["A2"].PutValue("Row 2 (hidden)");
                ws.Cells["A3"].PutValue("Row 3");

                // Hide the second row
                ws.Cells.SetRowHeight(1, ws.Cells.GetRowHeight(1));
                ws.Cells.Rows[1].IsHidden = true; // Correct way to hide a row

                wb.Save(path);
                Console.WriteLine($"Sample workbook created at '{path}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to create sample workbook: {ex.Message}");
                throw;
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                ExcludeHiddenRowsDemo.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}

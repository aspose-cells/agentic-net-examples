using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace BatchLoadExample
{
    // Custom filter that loads all data for visible sheets only,
    // otherwise loads only the structure.
    public class VisibleSheetLoadFilter : LoadFilter
    {
        public override void StartSheet(Worksheet sheet)
        {
            // Only load full data for visible worksheets.
            if (sheet.IsVisible)
            {
                // Load everything (cells, formulas, formatting, etc.).
                LoadDataFilterOptions = LoadDataFilterOptions.All;
            }
            else
            {
                // Load only the sheet structure (no cell data).
                LoadDataFilterOptions = LoadDataFilterOptions.Structure;
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Prepare a single LoadOptions instance with the custom filter.
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LoadFilter = new VisibleSheetLoadFilter();

            // List of workbook files to be loaded in batch.
            List<string> workbookFiles = new List<string>
            {
                "Template1.xlsx",
                "Template2.xlsx",
                "Template3.xlsx"
            };

            // Process each workbook using the same LoadOptions.
            foreach (string filePath in workbookFiles)
            {
                // Load the workbook with the shared LoadOptions.
                Workbook wb = new Workbook(filePath, loadOptions);

                // Example operation: display loaded worksheet names.
                Console.WriteLine($"Workbook: {filePath}");
                foreach (Worksheet ws in wb.Worksheets)
                {
                    Console.WriteLine($" - Sheet: {ws.Name}, Visible: {ws.IsVisible}, Cells Loaded: {ws.Cells.Count}");
                }

                // Save the workbook to a new file to demonstrate that it can be saved after loading.
                string outputPath = System.IO.Path.GetFileNameWithoutExtension(filePath) + "_Processed.xlsx";
                wb.Save(outputPath);
                Console.WriteLine($"Saved processed workbook as: {outputPath}");
            }
        }
    }
}
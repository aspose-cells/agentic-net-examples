using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class LoadWorkbookWithFilterDemo
    {
        static void Main()
        {
            // Path to the source Excel file
            string sourceFile = "sample.xlsx";

            // Create LoadOptions and assign a custom LoadFilter.
            // The filter loads only cell data and charts, skipping other objects.
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LoadFilter = new LoadFilter(
                LoadDataFilterOptions.CellData |   // Load cell values, formulas and formatting
                LoadDataFilterOptions.Chart       // Load chart objects
            );

            // Open the workbook with the specified load options.
            Workbook workbook = new Workbook(sourceFile, loadOptions);

            // Access the first worksheet to demonstrate that data was loaded.
            Worksheet sheet = workbook.Worksheets[0];
            Console.WriteLine($"Worksheet '{sheet.Name}' has {sheet.Cells.Count} cells loaded.");

            // List loaded charts (if any) in the worksheet.
            int chartCount = sheet.Charts.Count;
            Console.WriteLine($"Number of charts loaded: {chartCount}");
            for (int i = 0; i < chartCount; i++)
            {
                Console.WriteLine($"Chart {i + 1} name: {sheet.Charts[i].Name}");
            }

            // Save the workbook to verify that it was loaded correctly.
            string outputFile = "FilteredOutput.xlsx";
            workbook.Save(outputFile, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to '{outputFile}'.");
        }
    }
}
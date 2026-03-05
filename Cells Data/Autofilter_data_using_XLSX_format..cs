using System;
using Aspose.Cells;

namespace AsposeCellsAutoFilterDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file
            string inputPath = "input.xlsx";

            // Configure load options to enable auto‑filtering when the workbook is loaded
            LoadOptions loadOptions = new LoadOptions
            {
                AutoFilter = true
            };

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Access the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the range that contains the header row and data to be filtered
            // Example: columns A to D, rows 1 to 10 (Excel addresses are 1‑based)
            worksheet.AutoFilter.Range = "A1:D10";

            // Apply a filter on the first column (field index 0) for the value "Bananas"
            worksheet.AutoFilter.Filter(0, "Bananas");

            // Refresh the filter to hide rows that do not meet the criteria
            worksheet.AutoFilter.Refresh();

            // Save the filtered workbook to a new file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"AutoFilter applied and workbook saved to '{outputPath}'.");
        }
    }
}
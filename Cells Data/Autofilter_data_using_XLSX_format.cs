using System;
using Aspose.Cells;

namespace AutofilterDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file
            string inputPath = "data.xlsx";
            // Path for the filtered output XLSX file
            string outputPath = "filtered.xlsx";

            // Create LoadOptions and enable AutoFilter during loading
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.AutoFilter = true;

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(inputPath, loadOptions);
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the range that contains the header row and data to be filtered
            // Example assumes data is in columns A to D, rows 1 to 20
            worksheet.AutoFilter.Range = "A1:D20";

            // Apply a filter on the first column (index 0) for the value "Bananas"
            worksheet.AutoFilter.Filter(0, "Bananas");

            // Refresh the filter to hide rows that do not meet the criteria
            worksheet.AutoFilter.Refresh();

            // Save the filtered workbook in XLSX format
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }
}
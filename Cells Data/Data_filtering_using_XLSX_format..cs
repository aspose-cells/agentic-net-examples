using System;
using Aspose.Cells;

namespace AsposeCellsDataFiltering
{
    class Program
    {
        static void Main()
        {
            // Paths for the source workbook and the filtered result
            string sourcePath = "input.xlsx";
            string outputPath = "filtered_output.xlsx";

            // Create LoadOptions and enable AutoFilter during loading
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.AutoFilter = true; // AutoFilter will be applied when the file is opened

            // Load the workbook with the specified LoadOptions
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Get the first worksheet (you can change the index as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the range that contains the header row and data rows
            // Example range: columns A to C, rows 1 to 10
            worksheet.AutoFilter.Range = "A1:C10";

            // Apply a filter on the second column (index 1, column B) for the value "Electronics"
            worksheet.AutoFilter.Filter(1, "Electronics");

            // Refresh the filter to hide rows that do not match the criteria
            worksheet.AutoFilter.Refresh();

            // Save the filtered workbook to a new file
            workbook.Save(outputPath);
        }
    }
}
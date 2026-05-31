using System;
using Aspose.Cells;

namespace AsposeCellsLoadDataOnlyDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source workbook
            string sourcePath = "input.xlsx";

            // Create LoadOptions instance
            LoadOptions loadOptions = new LoadOptions();

            // Configure the LoadFilter to load only cell values (no formatting, formulas, etc.)
            // LoadDataFilterOptions.CellValue loads only the values of cells.
            loadOptions.LoadFilter = new LoadFilter(LoadDataFilterOptions.CellValue);

            // Load the workbook using the configured options
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Example: read a cell value to verify that data is loaded
            Worksheet sheet = workbook.Worksheets[0];
            Console.WriteLine("A1 value: " + sheet.Cells["A1"].StringValue);

            // Save the workbook (optional, demonstrates that it can be saved after loading)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);

            // Clean up
            workbook.Dispose();

            Console.WriteLine("Workbook loaded with cell values only and saved to: " + outputPath);
        }
    }
}
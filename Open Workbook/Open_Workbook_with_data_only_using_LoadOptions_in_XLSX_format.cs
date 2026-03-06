using System;
using Aspose.Cells;

namespace AsposeCellsLoadDataOnly
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file
            string sourcePath = "input.xlsx";

            // Create LoadOptions instance
            LoadOptions loadOptions = new LoadOptions();

            // Configure the LoadFilter to load only cell values (no formatting, charts, etc.)
            loadOptions.LoadFilter = new LoadFilter(LoadDataFilterOptions.CellValue);

            // Load the workbook using the constructor that accepts a file path and LoadOptions
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Demonstrate that data has been loaded
            Console.WriteLine("Number of worksheets loaded: " + workbook.Worksheets.Count);
            Console.WriteLine("Value of cell A1: " + workbook.Worksheets[0].Cells["A1"].StringValue);

            // Save the workbook to verify successful load (optional)
            workbook.Save("output.xlsx");
        }
    }
}
using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsLoadDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file
            string sourcePath = "input.xlsx";

            // Create LoadOptions and configure desired settings
            LoadOptions loadOptions = new LoadOptions();

            // Do not parse formulas during load (faster loading)
            loadOptions.ParsingFormulaOnOpen = false;

            // Load only cell values (skip formulas, formatting, etc.)
            loadOptions.LoadFilter = new LoadFilter(LoadDataFilterOptions.CellValue);

            // Set culture for proper number/date interpretation
            loadOptions.CultureInfo = new CultureInfo("en-US");

            // Load the workbook using the constructor that accepts a file path and LoadOptions
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Read and display the value of cell A1
            Cell cell = sheet.Cells["A1"];
            Console.WriteLine($"A1 value after load: {cell.StringValue}");

            // Modify a cell to demonstrate that the workbook is editable
            sheet.Cells["B2"].PutValue("Loaded with custom options");

            // Save the modified workbook to a new file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
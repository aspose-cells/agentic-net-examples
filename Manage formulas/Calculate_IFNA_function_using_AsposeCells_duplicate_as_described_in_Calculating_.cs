using System;
using Aspose.Cells;

namespace AsposeCellsIFNADemo
{
    class Program
    {
        static void Main()
        {
            // Path to the existing XLSX workbook
            string inputPath = "input.xlsx";

            // Load the workbook with default options (parsing formulas on open)
            LoadOptions loadOptions = new LoadOptions();
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Access the first worksheet and its cells collection
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Create a formula that generates an error (division by zero)
            cells["A1"].Formula = "=1/0";

            // Use IFNA to handle the possible error from A1
            // If A1 results in an error, B1 will display "No error"
            cells["B1"].Formula = "=IFNA(A1, \"No error\")";

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Output the calculated values to the console
            Console.WriteLine("A1 (error value): " + cells["A1"].StringValue);
            Console.WriteLine("B1 (IFNA result): " + cells["B1"].StringValue);

            // Save the workbook with the calculated results
            workbook.Save("output.xlsx");
        }
    }
}
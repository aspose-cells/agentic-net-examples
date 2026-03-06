using System;
using Aspose.Cells;

namespace AsposeCellsCellAccessDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing XLSX workbook (replace with your file path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath); // load rule

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Access the Cells collection
            Cells cells = worksheet.Cells;

            // ----- Access cells by zero‑based row and column indexes -----
            // Row 0, Column 0 corresponds to cell A1
            Cell cellA1 = cells[0, 0];
            Console.WriteLine($"A1 (by index) = {cellA1.StringValue}");

            // Row 1, Column 2 corresponds to cell C2
            Cell cellC2 = cells[1, 2];
            Console.WriteLine($"C2 (by index) = {cellC2.StringValue}");

            // ----- Access cells by Excel cell name -----
            Cell cellB3 = cells["B3"];
            Console.WriteLine($"B3 (by name) = {cellB3.StringValue}");

            // ----- Modify a cell value using the indexer -----
            cells[2, 1].PutValue("Updated Value"); // updates cell B3 (row 2, column 1)

            // ----- Read a numeric value and perform a calculation -----
            Cell numericCell = cells["D5"];
            if (numericCell != null && numericCell.Type == CellValueType.IsNumeric)
            {
                double original = numericCell.DoubleValue;
                double doubled = original * 2;
                cells["E5"].PutValue(doubled);
                Console.WriteLine($"D5 = {original}, E5 (doubled) = {doubled}");
            }

            // Save the workbook with changes (replace with your desired output path)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath); // save rule

            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
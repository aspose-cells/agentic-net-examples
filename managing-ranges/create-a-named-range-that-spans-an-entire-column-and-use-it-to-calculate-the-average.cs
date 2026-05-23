using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsNamedRangeAverage
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate some numeric data in column B (index 1)
                cells["B1"].PutValue(10);
                cells["B2"].PutValue(20);
                cells["B3"].PutValue(30);
                cells["B4"].PutValue(40);
                cells["B5"].PutValue(50);

                // Create a range that starts at B1 (single cell)
                // Using CreateRange(int firstRow, int firstColumn, int totalRows, int totalColumns)
                AsposeRange singleCellRange = cells.CreateRange(0, 1, 1, 1);

                // Expand the range to the entire column that contains the single cell
                // Using Range.EntireColumn property
                AsposeRange entireColumnRange = singleCellRange.EntireColumn;

                // Assign a name to the entire column range
                // The Name property of Range sets the defined name
                entireColumnRange.Name = "MyColumn";

                // Use the named range in a formula to calculate the average
                cells["C1"].Formula = "=AVERAGE(MyColumn)";

                // Calculate all formulas in the workbook
                workbook.CalculateFormula();

                // Retrieve and display the calculated average
                Console.WriteLine("Average of MyColumn: " + cells["C1"].DoubleValue);

                // Save the workbook (lifecycle rule: save)
                string outputPath = "NamedRangeAverage.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
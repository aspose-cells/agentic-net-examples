using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

class SparklineValidationDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data, including a formula that will cause an error (#DIV/0!)
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(0);
            sheet.Cells["C1"].Formula = "=A1/B1"; // division by zero error
            sheet.Cells["D1"].PutValue(3);

            // Define the data range for the sparkline
            string dataRange = "A1:D1";

            // Recalculate formulas so that error values are materialized
            workbook.CalculateFormula();

            // Validate that the data range does not contain any error cells
            bool hasError = false;
            string[] rangeParts = dataRange.Split(':');
            CellArea rangeArea = CellArea.CreateCellArea(rangeParts[0], rangeParts[1]);

            for (int row = rangeArea.StartRow; row <= rangeArea.EndRow && !hasError; row++)
            {
                for (int col = rangeArea.StartColumn; col <= rangeArea.EndColumn; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    if (cell.Type == CellValueType.IsError)
                    {
                        hasError = true;
                        Console.WriteLine($"Error detected in cell {cell.Name}");
                        break;
                    }
                }
            }

            if (!hasError)
            {
                // Data range is clean – create the sparkline group
                CellArea sparklineLocation = CellArea.CreateCellArea("E1", "E1");
                int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, dataRange, false, sparklineLocation);
                SparklineGroup group = sheet.SparklineGroups[groupIndex];
                Console.WriteLine("Sparkline created successfully.");
            }
            else
            {
                Console.WriteLine("Sparkline creation skipped due to errors in the data range.");
            }

            // Save the workbook
            string outputPath = "SparklineValidationDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
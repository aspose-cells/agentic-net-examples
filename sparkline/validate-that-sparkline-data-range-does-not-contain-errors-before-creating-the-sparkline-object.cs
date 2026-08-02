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

            // Populate sample data (including an intentional error)
            sheet.Cells["A1"].PutValue(5);
            // Use the Formula property to set a formula (PutFormula may not be available in some versions)
            sheet.Cells["A2"].Formula = "=1/0"; // This will produce a #DIV/0! error
            sheet.Cells["A3"].PutValue(7);
            sheet.Cells["A4"].PutValue(2);
            sheet.Cells["A5"].PutValue(9);

            // Define the data range for the sparkline
            string dataRange = "A1:A5";

            // Validate that the data range does not contain error cells
            if (ContainsError(sheet, dataRange))
            {
                Console.WriteLine("Data range contains error cells. Sparkline will not be created.");
            }
            else
            {
                // Define where the sparkline will be placed
                CellArea location = new CellArea
                {
                    StartRow = 0,
                    EndRow = 0,
                    StartColumn = 1,
                    EndColumn = 1
                };

                // Add a sparkline group using the validated data range
                int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, dataRange, false, location);
                SparklineGroup group = sheet.SparklineGroups[groupIndex];

                Console.WriteLine("Sparkline created successfully.");
            }

            // Save the workbook
            string outputPath = "SparklineValidationDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Checks each cell in the specified range for an error value
    static bool ContainsError(Worksheet sheet, string range)
    {
        // Convert the range string to a CellArea
        string[] parts = range.Split(':');
        CellArea area = CellArea.CreateCellArea(parts[0], parts[1]);

        for (int row = area.StartRow; row <= area.EndRow; row++)
        {
            for (int col = area.StartColumn; col <= area.EndColumn; col++)
            {
                Cell cell = sheet.Cells[row, col];
                if (cell.Type == CellValueType.IsError)
                {
                    return true;
                }
            }
        }
        return false;
    }
}
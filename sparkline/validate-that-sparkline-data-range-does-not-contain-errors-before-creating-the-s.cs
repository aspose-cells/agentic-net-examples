using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSparklineValidation
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the sparkline
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(3);
            sheet.Cells["C1"].PutValue(7);
            sheet.Cells["D1"].PutValue(2);

            // Define the data range for the sparkline
            string dataRange = "A1:D1";

            // Convert the range string to a CellArea for validation
            CellArea rangeArea = CellArea.CreateCellArea("A1", "D1");

            // Validate that the data range does not contain error cells
            bool hasError = false;
            for (int row = rangeArea.StartRow; row <= rangeArea.EndRow && !hasError; row++)
            {
                for (int col = rangeArea.StartColumn; col <= rangeArea.EndColumn; col++)
                {
                    // Check if the cell contains an error value
                    if (sheet.Cells[row, col].Type == CellValueType.IsError)
                    {
                        hasError = true;
                        break;
                    }
                }
            }

            if (!hasError)
            {
                // Define where the sparkline will be placed (cell E1)
                CellArea location = new CellArea
                {
                    StartRow = 0,
                    EndRow = 0,
                    StartColumn = 4,
                    EndColumn = 4
                };

                // Add a sparkline group with the validated data range
                int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, dataRange, false, location);
                SparklineGroup group = sheet.SparklineGroups[groupIndex];

                // Optionally add a sparkline to the group (redundant when using Add with parameters)
                // group.Sparklines.Add(sheet.Name + "!" + dataRange, 0, 4);
            }
            else
            {
                Console.WriteLine("The specified data range contains error cells. Sparkline not created.");
            }

            // Save the workbook
            workbook.Save("ValidatedSparkline.xlsx");
        }
    }
}
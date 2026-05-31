using System;
using System.Collections.Generic;
using Aspose.Cells;

class ThresholdReport
{
    static void Main()
    {
        // Load the source workbook (replace with actual file path)
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet sourceSheet = workbook.Worksheets[0];

        // Define the numeric threshold
        double threshold = 100.0;

        // List to hold addresses of cells that exceed the threshold
        List<string> violatingAddresses = new List<string>();

        // Determine the used range of the worksheet
        int maxRow = sourceSheet.Cells.MaxDataRow;
        int maxCol = sourceSheet.Cells.MaxDataColumn;

        // Scan each cell within the used range
        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                Cell cell = sourceSheet.Cells[row, col];

                // Check if the cell contains a numeric value greater than the threshold
                if (cell.Type == CellValueType.IsNumeric && cell.DoubleValue > threshold)
                {
                    violatingAddresses.Add(cell.Name);
                }
            }
        }

        // Create a new worksheet to hold the report
        int reportIndex = workbook.Worksheets.Add();
        Worksheet reportSheet = workbook.Worksheets[reportIndex];
        reportSheet.Name = "ThresholdReport";

        // Write header row
        reportSheet.Cells[0, 0].PutValue("Cell Address");
        reportSheet.Cells[0, 1].PutValue("Value");

        // Populate the report with violating cell information
        for (int i = 0; i < violatingAddresses.Count; i++)
        {
            string address = violatingAddresses[i];
            Cell srcCell = sourceSheet.Cells[address];

            reportSheet.Cells[i + 1, 0].PutValue(address);
            reportSheet.Cells[i + 1, 1].PutValue(srcCell.DoubleValue);
        }

        // Save the workbook with the report
        workbook.Save("output.xlsx");
    }
}
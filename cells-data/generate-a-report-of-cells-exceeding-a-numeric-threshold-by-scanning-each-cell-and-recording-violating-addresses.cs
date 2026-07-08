using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

class ThresholdReport
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Work with the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Define the numeric threshold
        double threshold = 100.0;

        // List to store addresses of cells that exceed the threshold
        List<string> violatingAddresses = new List<string>();

        // Enumerate all cells in the worksheet
        IEnumerator enumerator = worksheet.Cells.GetEnumerator();
        while (enumerator.MoveNext())
        {
            Cell cell = (Cell)enumerator.Current;

            // Check if the cell contains a numeric value and exceeds the threshold
            if (cell.IsNumericValue && cell.DoubleValue > threshold)
            {
                violatingAddresses.Add(cell.Name); // Store the cell address (e.g., "B5")
            }
        }

        // Create a new worksheet to hold the report
        int reportIndex = workbook.Worksheets.Add();
        Worksheet reportSheet = workbook.Worksheets[reportIndex];
        reportSheet.Name = "ThresholdReport";

        // Write header
        reportSheet.Cells[0, 0].PutValue($"Cells exceeding threshold {threshold}");

        // Write each violating address into the report sheet
        for (int i = 0; i < violatingAddresses.Count; i++)
        {
            reportSheet.Cells[i + 1, 0].PutValue(violatingAddresses[i]);
        }

        // Save the workbook with the report
        workbook.Save("output.xlsx");
    }
}
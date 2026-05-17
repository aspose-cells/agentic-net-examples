using System;
using Aspose.Cells;

namespace AsposeCellsDebugDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (using the provided creation rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Sample merged cells to demonstrate iteration
            // Merge A1:B2 (rows 0-1, columns 0-1)
            cells.Merge(0, 0, 1, 1);
            // Merge D4:E5 (rows 3-4, columns 3-4)
            cells.Merge(3, 3, 1, 1);

            // Retrieve all merged cell areas using the GetMergedAreas method (rule exists)
            CellArea[] mergedAreas = cells.GetMergedAreas();

            // Iterate through each CellArea and log its row/column indices for debugging
            foreach (CellArea area in mergedAreas)
            {
                Console.WriteLine(
                    $"Merged Area -> StartRow: {area.StartRow}, StartColumn: {area.StartColumn}, " +
                    $"EndRow: {area.EndRow}, EndColumn: {area.EndColumn}");
            }

            // If there are other mapped cell areas (e.g., validation areas), they can be processed similarly:
            // Example: iterate through validation areas if any exist
            if (worksheet.Validations.Count > 0)
            {
                foreach (Validation validation in worksheet.Validations)
                {
                    foreach (CellArea vArea in validation.Areas)
                    {
                        Console.WriteLine(
                            $"Validation Area -> StartRow: {vArea.StartRow}, StartColumn: {vArea.StartColumn}, " +
                            $"EndRow: {vArea.EndRow}, EndColumn: {vArea.EndColumn}");
                    }
                }
            }

            // Save the workbook (using the provided save rule)
            workbook.Save("DebugDemo.xlsx");
        }
    }
}
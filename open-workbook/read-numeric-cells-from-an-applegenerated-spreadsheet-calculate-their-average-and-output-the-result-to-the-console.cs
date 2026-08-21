// Title: Compute the average of numeric cells in an Apple Numbers‑generated XLSX with Aspose.Cells for .NET
// Description: Loads an Apple Numbers‑exported workbook using Aspose.Cells, scans the used range of the first worksheet, sums cells flagged as IsNumericValue, calculates the mean, and writes the result to the console.
// Keywords: Aspose.Cells .NET average numeric cells | read numeric values Apple Numbers XLSX | C# iterate used range Excel | calculate spreadsheet mean Aspose | console utility Excel statistics | Apple Numbers export processing
// Common Searches: Aspose.Cells calculate average of numbers in Excel file | C# read only numeric cells from Apple Numbers export | How to get mean of all values in a workbook using Aspose.Cells | Iterate used range and sum numeric cells Aspose .NET | Console program to average numeric cells in XLSX
// Developer Intent: Determine the mean value of every numeric cell in a loaded workbook and display it.
// Use Cases: Produce quick summary metrics for financial reports exported from Apple Numbers. | Validate data quality across multiple macOS‑generated spreadsheets by checking the overall average. | Build a lightweight command‑line tool that flags unusually high or low averages in engineering data sets. | Integrate into automated pipelines that need a numeric‑only checksum before further processing.
// AI Prompts: Create a reusable function that accepts an Aspose.Cells Workbook and returns the average of its numeric cells, excluding dates. | Modify the example to handle multiple worksheets and output each sheet's average separately. | Add robust error handling for missing files, empty worksheets, and non‑numeric content while logging detailed diagnostics.

using System;
using Aspose.Cells;

// Loads an Apple Numbers‑exported workbook using Aspose.Cells, scans the used range of the first worksheet, sums cells flagged as IsNumericValue, calculates the mean, and writes the result to the console.
class Program
{
    static void Main()
    {
        // Path to the Apple‑generated spreadsheet (replace with actual file path)
        string filePath = "input.xlsx";

        // Load the workbook
        Workbook workbook = new Workbook(filePath);
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        double sum = 0;
        int count = 0;

        // Determine the used range of the worksheet
        int maxRow = cells.MaxDataRow;
        int maxCol = cells.MaxDataColumn;

        // Iterate through all cells in the used range
        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                Cell cell = cells[row, col];
                // Check if the cell contains a numeric value (including dates and times)
                if (cell != null && cell.IsNumericValue)
                {
                    sum += cell.DoubleValue;
                    count++;
                }
            }
        }

        // Calculate the average
        double average = count > 0 ? sum / count : 0;

        // Output the result to the console
        Console.WriteLine($"Average of numeric cells: {average}");
    }
}

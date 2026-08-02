// Title: Compute Average of All Numeric Cells in an Apple‑Generated XLSX with Aspose.Cells for .NET (C#)
// Description: Loads an Apple‑generated XLSX file using Aspose.Cells, converts numeric strings to true numbers, iterates through every cell in the first worksheet, sums values that are numeric, calculates the average, and writes the result to the console.
// Keywords: Aspose.Cells | C# average numeric cells | read Excel numeric values | convert string to numeric Aspose.Cells | iterate worksheet cells | calculate worksheet average | Apple generated spreadsheet
// Common Searches: Aspose.Cells calculate average of numeric cells | C# read numeric values from Excel and compute average | How to sum numeric cells using Aspose.Cells | Convert numeric strings to numbers with Aspose.Cells | Average of all cells in first worksheet Aspose
// Developer Intent: Calculate the average of every numeric cell in the first worksheet of an Apple‑generated XLSX and display the result.
// Use Cases: Create a quick statistical summary (average) of numeric data across a sheet for reporting. | Validate data quality after importing spreadsheets from external sources by checking average values. | Perform lightweight data analysis without Excel formulas, using Aspose.Cells to process large workbooks programmatically.
// AI Prompts: Generate C# code that uses Aspose.Cells to compute the average of numeric cells in a workbook, including conversion of numeric strings. | Suggest an efficient way to iterate only the used range (MaxRow/MaxColumn) when summing numeric values with Aspose.Cells. | Explain how to exclude date values from the average calculation while still using the IsNumericValue property.

using System;
using Aspose.Cells;

// Loads an Apple‑generated XLSX file using Aspose.Cells, converts numeric strings to true numbers, iterates through every cell in the first worksheet, sums values that are numeric, calculates the average, and writes the result to the console.
class Program
{
    static void Main()
    {
        // Path to the Apple‑generated spreadsheet
        string inputPath = "input.xlsx";

        // Load the workbook (creation/load rule)
        Workbook workbook = new Workbook(inputPath);

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Convert any numeric strings to actual numbers (optional but useful)
        cells.ConvertStringToNumericValue();

        double sum = 0;
        int count = 0;

        // Enumerate all cells in the worksheet
        foreach (Cell cell in cells)
        {
            // Consider only cells that contain a numeric value (int, double, date, etc.)
            if (cell.IsNumericValue)
            {
                sum += cell.DoubleValue;
                count++;
            }
        }

        double average = count > 0 ? sum / count : 0;

        // Output the result to the console
        Console.WriteLine($"Average of numeric cells: {average}");
    }
}

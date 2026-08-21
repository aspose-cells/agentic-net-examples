// Title: Convert string‑based numbers to true numeric values across all worksheets with Aspose.Cells for .NET
// Description: Loads an Excel workbook, runs Cells.ConvertStringToNumericValue on every worksheet to turn text‑stored numbers into real numeric types, and saves the updated file.
// Keywords: Aspose.Cells | ConvertStringToNumericValue | C# Excel numeric conversion | string to number Excel .NET | batch worksheet conversion | Excel workbook data cleanup | numeric type casting Aspose
// Common Searches: Aspose.Cells convert text numbers to numeric in all sheets | C# example for Cells.ConvertStringToNumericValue workbook | how to change string numbers to numbers in Excel using Aspose | batch convert numeric strings in Excel with .NET | convert string based values to numbers Aspose.Cells
// Developer Intent: Transform every text‑based numeric cell in a workbook into a proper numeric value.
// Use Cases: Standardize imported CSV data where numbers appear as text before calculations. | Prepare a multi‑sheet report for charting or formula evaluation by ensuring numeric cells are recognized. | Automate data cleansing in a CI pipeline that processes Excel files from various sources.
// AI Prompts: Show how to limit ConvertStringToNumericValue to a specific range instead of the whole sheet. | Explain how to retrieve a list of cells that were changed from text to numeric after conversion. | Provide code that logs the address and original value of each cell converted by ConvertStringToNumericValue.

using System;
using Aspose.Cells;

// Loads an Excel workbook, runs Cells.ConvertStringToNumericValue on every worksheet to turn text‑stored numbers into real numeric types, and saves the updated file.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Convert string-based numeric values to true numbers in every worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.Cells.ConvertStringToNumericValue();
        }

        // Save the modified workbook
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
    }
}

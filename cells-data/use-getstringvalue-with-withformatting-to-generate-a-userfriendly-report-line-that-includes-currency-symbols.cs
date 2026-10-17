// Title: Generate a currency‑formatted report line using GetStringValue with DisplayString in Aspose.Cells for .NET
// AI Prompts: Show how to apply a built‑in currency number format to a cell and retrieve the formatted text with GetStringValue(DisplayString) in C#. | Demonstrate building a sales summary string that includes the currency symbol by reading a formatted cell value from an Aspose.Cells workbook. | Explain how to combine cell styling and CellValueFormatStrategy to produce a user‑friendly report line in a .NET console app.
// Common Searches: Aspose.Cells C# GetStringValue DisplayString currency symbol example | How to read formatted currency value from Excel cell using Aspose.Cells | Create report line with currency formatting in Aspose.Cells .NET | Apply built‑in currency format and get formatted string Aspose.Cells | CellValueFormatStrategy DisplayString usage for monetary values in C#
// Tags: GetStringValue with DisplayString for currency | apply built‑in currency number format Aspose.Cells | retrieve formatted monetary value from Excel cell | C# report line generation using cell formatting | Aspose.Cells workbook save with formatted data

using Aspose.Cells;
using System;

// The example creates a workbook, writes a numeric value to cell A1, applies a built‑in currency format, retrieves the value as a formatted string with GetStringValue(DisplayString), composes a report line that includes the currency symbol, prints it, and saves the workbook as Report.xlsx.
class ReportGenerator
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Put a numeric value into cell A1
            Cell cell = sheet.Cells["A1"];
            cell.PutValue(1234.56);

            // Apply a built‑in currency format (e.g., $#,##0.00)
            Style style = cell.GetStyle();
            style.Number = 164; // Built‑in currency format
            cell.SetStyle(style);

            // Retrieve the cell value as a formatted string (includes currency symbol)
            string reportLine = $"Total sales: {cell.GetStringValue(CellValueFormatStrategy.DisplayString)}";

            // Output the report line
            Console.WriteLine(reportLine);

            // Save the workbook (lifecycle rule)
            workbook.Save("Report.xlsx");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}

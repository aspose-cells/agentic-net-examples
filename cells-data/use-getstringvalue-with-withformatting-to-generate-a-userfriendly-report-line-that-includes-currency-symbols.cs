// Title: Retrieve a currency‑formatted string with GetStringValue (DisplayString) in Aspose.Cells for .NET
// Description: Demonstrates how to place a numeric value in a worksheet, apply the built‑in currency style (Number format ID 4), extract the formatted text using GetStringValue with the DisplayString strategy, and compose a readable report line that includes the currency symbol. The example also shows writing the line to other cells and saving the workbook.
// Keywords: Aspose.Cells GetStringValue | DisplayString strategy | currency format .NET | CellValueFormatStrategy example | formatted cell value report | C# Aspose.Cells currency | Excel currency symbol extraction
// Common Searches: Aspose.Cells get cell value with currency symbol | GetStringValue DisplayString C# example | How to format numbers as currency in Aspose.Cells | Create report line from formatted Excel cell | Retrieve formatted value from worksheet cell
// Developer Intent: Extract a cell’s currency‑formatted text and embed it in a human‑readable report string.
// Use Cases: Generate a sales summary line that shows the amount with the correct currency symbol. | Insert a formatted monetary value into another cell for documentation or further calculations. | Save a workbook after adding a user‑friendly report line that reflects cell styling.
// AI Prompts: Show how to use Aspose.Cells GetStringValue with CellValueFormatStrategy.DisplayString to obtain a currency‑formatted string in C#. | Provide a C# snippet that applies the built‑in currency style to a cell, reads the formatted value, and writes a report line to another cell. | Explain combining GetStringValue, formatting, and string interpolation to create a readable sales report in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsReportDemo
{
    // Demonstrates how to place a numeric value in a worksheet, apply the built‑in currency style (Number format ID 4), extract the formatted text using GetStringValue with the DisplayString strategy, and compose a readable report line that includes the currency symbol. The example also shows writing the line to other cells and saving the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Target cell that will hold the numeric value
            Cell amountCell = sheet.Cells["B2"];
            // Put a numeric value (e.g., total sales)
            amountCell.PutValue(1234.56);

            // Apply a built‑in currency format (Number format ID 4 = "$#,##0.00")
            Style currencyStyle = amountCell.GetStyle();
            currencyStyle.Number = 4; // Currency format
            amountCell.SetStyle(currencyStyle);

            // Retrieve the formatted string using GetStringValue with DisplayString strategy
            // This includes the currency symbol and respects the cell's style.
            string formattedAmount = amountCell.GetStringValue(CellValueFormatStrategy.DisplayString);

            // Build a user‑friendly report line
            string reportLine = $"Total Sales: {formattedAmount}";

            // Output the report line to the console
            Console.WriteLine(reportLine);

            // Optionally, write the report line into another cell for demonstration
            sheet.Cells["A2"].PutValue("Report:");
            sheet.Cells["A3"].PutValue(reportLine);

            // Save the workbook (lifecycle rule: use provided save method)
            workbook.Save("ReportDemo.xlsx");
        }
    }
}

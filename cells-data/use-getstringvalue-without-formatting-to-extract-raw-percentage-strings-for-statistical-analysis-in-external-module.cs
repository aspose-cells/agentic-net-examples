// Title: Get raw percentage string from a formatted cell using GetStringValue (CellValueFormatStrategy.None) in Aspose.Cells for .NET
// Description: Demonstrates how to write a decimal (0.25) to a cell, apply the built‑in percentage format, and then retrieve the underlying raw string "0.25" with GetStringValue(CellValueFormatStrategy.None). The example also shows the formatted output ("25%") and saves the workbook, enabling statistical analysis without percent symbols.
// Keywords: Aspose.Cells GetStringValue | CellValueFormatStrategy.None | percentage format raw value | C# unformatted cell string | extract numeric string Aspose.Cells | raw decimal from formatted cell | Aspose.Cells .NET example
// Common Searches: Aspose.Cells get raw value from percentage cell | GetStringValue without formatting C# | CellValueFormatStrategy.None percentage example | retrieve underlying numeric string Aspose.Cells | how to ignore number format in Aspose.Cells
// Developer Intent: Retrieve the unformatted numeric string of a percentage‑formatted cell for further processing or analysis.
// Use Cases: Read decimal values from cells displayed as percentages and feed them into statistical models. | Validate that number‑formatting rules are applied correctly by comparing raw and formatted outputs. | Export data to CSV or JSON while preserving original numeric values, not the displayed format.
// AI Prompts: Show how to use GetStringValue with CellValueFormatStrategy.None to extract raw values from percentage‑formatted cells in Aspose.Cells for .NET. | Provide a C# snippet that reads a range of percentage cells and stores their underlying decimals in a list without applying any formatting. | Explain the difference between StringValue and GetStringValue(CellValueFormatStrategy.None) when handling various number formats in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    // Demonstrates how to write a decimal (0.25) to a cell, apply the built‑in percentage format, and then retrieve the underlying raw string "0.25" with GetStringValue(CellValueFormatStrategy.None). The example also shows the formatted output ("25%") and saves the workbook, enabling statistical analysis without percent symbols.
    public class GetRawPercentageDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Put a numeric value that represents 25%
                worksheet.Cells["A1"].PutValue(0.25);

                // Apply the built‑in percentage number format (9) to the cell
                Style percentStyle = worksheet.Cells["A1"].GetStyle();
                percentStyle.Number = 9; // Percentage format
                worksheet.Cells["A1"].SetStyle(percentStyle);

                // Retrieve the cell value as a raw string without any formatting
                string rawValue = worksheet.Cells["A1"].GetStringValue(CellValueFormatStrategy.None);
                Console.WriteLine("Raw value (no format): " + rawValue); // Expected output: 0.25

                // For comparison, retrieve the formatted string (e.g., "25%")
                string formattedValue = worksheet.Cells["A1"].StringValue;
                Console.WriteLine("Formatted value: " + formattedValue); // Expected output: 25%

                // Save the workbook (optional, demonstrates lifecycle usage)
                workbook.Save("GetRawPercentageDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("Error: " + ex.Message);
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            GetRawPercentageDemo.Run();
        }
    }
}

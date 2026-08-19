// Title: Extract raw percentage strings from formatted cells with Aspose.Cells GetStringValue (No formatting) – C# .NET
// Description: Demonstrates how to create a workbook, apply the built‑in percent style to cells, and use GetStringValue(CellValueFormatStrategy.None) to retrieve the underlying numeric strings (e.g., "0.5", "0.75", "0.33") without any formatting, then optionally save the file.
// Keywords: Aspose.Cells GetStringValue | CellValueFormatStrategy.None | raw percentage value C# | unformatted cell value .NET | extract numeric string from percent cell | Aspose.Cells example | C# spreadsheet data extraction
// Common Searches: Aspose.Cells get raw value from percent formatted cell | CellValueFormatStrategy.None usage example | retrieve underlying numeric string Aspose.Cells C# | how to ignore cell formatting when reading values | extract unformatted data from Excel with Aspose
// Developer Intent: Obtain the exact numeric strings stored in cells that are displayed as percentages, bypassing any applied number format.
// Use Cases: Feed raw percentage numbers into statistical or machine‑learning models. | Export unformatted numeric data to CSV/JSON for downstream analytics. | Validate source values before applying custom business rules or re‑formatting.
// AI Prompts: Write a C# loop that reads a range of percent‑formatted cells and stores each raw numeric string using GetStringValue(CellValueFormatStrategy.None) with Aspose.Cells. | Compare CellValueFormatStrategy.None, Displayed, and FormulaResult, and advise when each should be used for data extraction. | Generate code that scans an entire worksheet, extracts raw values without formatting, and writes the results to a JSON file.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, apply the built‑in percent style to cells, and use GetStringValue(CellValueFormatStrategy.None) to retrieve the underlying numeric strings (e.g., "0.5", "0.75", "0.33") without any formatting, then optionally save the file.
    public class GetRawPercentageStringsDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Cell A1: numeric value 0.5 formatted as percent
            // -------------------------------------------------
            Cell cellA1 = worksheet.Cells["A1"];
            cellA1.PutValue(0.5);                     // underlying value is 0.5
            Style percentStyle = cellA1.GetStyle();
            percentStyle.Number = 10;                 // built‑in percent format (e.g., "0%")
            cellA1.SetStyle(percentStyle);

            // Retrieve the raw string without any formatting
            string rawA1 = cellA1.GetStringValue(CellValueFormatStrategy.None);
            Console.WriteLine($"A1 raw value (no format): {rawA1}");   // Expected: "0.5"

            // -------------------------------------------------
            // Cell A2: numeric value 0.75 formatted as percent
            // -------------------------------------------------
            Cell cellA2 = worksheet.Cells["A2"];
            cellA2.PutValue(0.75);
            Style percentStyle2 = cellA2.GetStyle();
            percentStyle2.Number = 10;                // percent format
            cellA2.SetStyle(percentStyle2);

            string rawA2 = cellA2.GetStringValue(CellValueFormatStrategy.None);
            Console.WriteLine($"A2 raw value (no format): {rawA2}");   // Expected: "0.75"

            // -------------------------------------------------
            // Cell A3: numeric value 0.33 without percent format
            // -------------------------------------------------
            Cell cellA3 = worksheet.Cells["A3"];
            cellA3.PutValue(0.33);
            // No special style applied; default is general

            string rawA3 = cellA3.GetStringValue(CellValueFormatStrategy.None);
            Console.WriteLine($"A3 raw value (no format): {rawA3}");   // Expected: "0.33"

            // -------------------------------------------------
            // Save the workbook (optional, just to demonstrate lifecycle)
            // -------------------------------------------------
            workbook.Save("RawPercentageDemo.xlsx");
        }
    }
}

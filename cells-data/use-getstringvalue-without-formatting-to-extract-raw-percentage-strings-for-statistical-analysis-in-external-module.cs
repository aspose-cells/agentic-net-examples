// Title: Extract raw numeric string from a percent‑formatted cell using GetStringValue with CellValueFormatStrategy.None in Aspose.Cells (.NET)
// AI Prompts: Write C# code that calls Cell.GetStringValue(CellValueFormatStrategy.None) to obtain the underlying numeric string of a cell formatted as a percentage. | Show how to retrieve the unformatted value "0.25" from a cell displaying "25%" in Aspose.Cells for statistical analysis. | Demonstrate extracting raw percentage values without formatting for further processing in an external module using Aspose.Cells.
// Common Searches: Aspose.Cells C# get unformatted value from percent formatted cell | Cell.GetStringValue with CellValueFormatStrategy.None example | how to read raw numeric string of a percentage cell in Aspose.Cells | extract underlying value from Excel percent format using Aspose.Cells .NET | retrieve raw data for statistical analysis from formatted Excel cell Aspose
// Tags: GetStringValue raw value extraction Aspose.Cells | CellValueFormatStrategy.None usage | percentage format raw string C# | unformatted cell value Aspose.Cells | statistical analysis raw Excel data Aspose

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example creates a workbook, writes the numeric value 0.25 to cell A1, applies Excel's built‑in percent format, then uses GetStringValue with CellValueFormatStrategy.None to obtain the raw string "0.25" while StringValue returns the formatted "25%", and finally saves the workbook as RawPercentageDemo.xlsx.
    class GetRawPercentageStrings
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Put a numeric value that represents 25%
            Cell percentCell = cells["A1"];
            percentCell.PutValue(0.25);

            // Apply Excel's built‑in percentage format (Number = 10)
            Style percentStyle = percentCell.GetStyle();
            percentStyle.Number = 10; // 10 = built‑in percent format
            percentCell.SetStyle(percentStyle);

            // Retrieve the raw value as a string without any formatting
            // CellValueFormatStrategy.None returns the underlying value unchanged
            string rawValue = percentCell.GetStringValue(CellValueFormatStrategy.None);
            Console.WriteLine("Raw percentage string (no formatting): " + rawValue);
            // Expected output: "0.25"

            // For comparison, show the formatted string that Excel would display
            string formattedValue = percentCell.StringValue;
            Console.WriteLine("Formatted percentage string: " + formattedValue);
            // Expected output: "25%"

            // Save the workbook (lifecycle: save)
            workbook.Save("RawPercentageDemo.xlsx");
        }
    }
}

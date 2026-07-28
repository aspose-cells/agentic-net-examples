// Title: Aspose.Cells for .NET: Custom number format to display zeros as a dash
// Description: C# example that creates a workbook, defines the custom format "0;-0;\"-\"" (positive, negative, zero), applies it to cells A1:A5, writes sample data including zeros, and saves the file as CustomZeroFormatDemo.xlsx. Zero values appear as a dash, ideal for clean financial or reporting layouts.
// Keywords: Aspose.Cells | C# | .NET | custom number format | zero dash | Excel formatting | display zero as dash | format positive negative zero | financial report formatting | Excel dash for zero values
// Common Searches: Aspose.Cells custom format zero dash | C# Excel format zero as dash | How to show dash instead of 0 in Aspose.Cells | Custom number format string 0;-0;"-" Aspose | Apply number format to range Aspose.Cells .NET
// Developer Intent: Create an Excel workbook and apply a style that renders zero values as a dash using Aspose.Cells for .NET.
// Use Cases: Financial statements where zero amounts are displayed as a dash to improve readability. | Automated reporting dashboards that replace zero metrics with a dash for a cleaner look. | Batch processing of multiple worksheets where the same zero‑as‑dash formatting is applied programmatically.
// AI Prompts: Generate C# code with Aspose.Cells that defines a custom number format "0;-0;\"-\"" and applies it to a column of data. | Explain each part of the custom number format string for positive, negative, and zero values in Aspose.Cells. | Show how to programmatically apply a zero‑as‑dash style to an entire worksheet range using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace Example
{
    // C# example that creates a workbook, defines the custom format "0;-0;\"-\"" (positive, negative, zero), applies it to cells A1:A5, writes sample data including zeros, and saves the file as CustomZeroFormatDemo.xlsx. Zero values appear as a dash, ideal for clean financial or reporting layouts.
    public class CustomZeroFormatDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Define a custom number format:
                // Positive numbers -> 0
                // Negative numbers -> -0
                // Zero values        -> dash ("-")
                Style zeroDashStyle = workbook.CreateStyle();
                zeroDashStyle.Custom = "0;-0;\"-\"";

                // Populate some sample data including zeros
                sheet.Cells["A1"].PutValue(123);
                sheet.Cells["A2"].PutValue(-45);
                sheet.Cells["A3"].PutValue(0);
                sheet.Cells["A4"].PutValue(78.9);
                sheet.Cells["A5"].PutValue(0);

                // Apply the custom style to the range A1:A5
                for (int row = 0; row < 5; row++)
                {
                    Cell cell = sheet.Cells[row, 0];
                    cell.SetStyle(zeroDashStyle);
                }

                // Save the workbook
                workbook.Save("CustomZeroFormatDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            CustomZeroFormatDemo.Run();
        }
    }
}

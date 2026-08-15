// Title: Aspose.Cells for .NET – C# custom number format that displays zero as a dash
// Description: Demonstrates how to create a workbook, define a style with the custom format "0;-0;\"-\"" to show positive numbers normally, negative numbers with a minus sign, and zero values as a dash, apply the number‑format part to a specific range using StyleFlag, populate sample data, and save the file as CustomZeroFormat.xlsx.
// Keywords: Aspose.Cells custom number format | C# zero dash format | display zero as dash | StyleFlag number format Aspose.Cells | apply format to range Aspose.Cells | Aspose.Cells for .NET formatting | Excel zero dash style
// Common Searches: Aspose.Cells how to format zero as dash | custom number format string 0;-0;"-" in C# | apply number format to specific cells Aspose.Cells | StyleFlag usage for number format Aspose.Cells | C# Excel zero dash formatting example
// Developer Intent: The developer needs to show zero values as a dash while preserving normal formatting for positive and negative numbers in a selected cell range.
// Use Cases: Financial reports where zero amounts are represented by a dash for cleaner presentation. | Inventory sheets that replace 0 quantities with a dash to indicate no stock. | Selective styling of a column or range without affecting other cell formats.
// AI Prompts: Show how to modify the custom format to display "N/A" instead of a dash in Aspose.Cells. | Provide code that applies the same zero‑dash format to an entire column using a loop. | Explain how to combine the zero‑dash number format with a date format in a single style.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsCustomZeroFormat
{
    // Demonstrates how to create a workbook, define a style with the custom format "0;-0;\"-\"" to show positive numbers normally, negative numbers with a minus sign, and zero values as a dash, apply the number‑format part to a specific range using StyleFlag, populate sample data, and save the file as CustomZeroFormat.xlsx.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Create a style with a custom number format:
                // Positive numbers: 0
                // Negative numbers: -0
                // Zero values: displayed as a dash ("-")
                Style style = workbook.CreateStyle();
                style.Custom = "0;-0;\"-\"";

                // Apply only the number format part of the style to the desired range
                StyleFlag flag = new StyleFlag();
                flag.NumberFormat = true;

                // Define a range (A1:A5) and apply the style
                AsposeRange range = sheet.Cells.CreateRange("A1", "A5");
                range.ApplyStyle(style, flag);

                // Populate the cells with sample values
                sheet.Cells["A1"].PutValue(123);   // Positive number
                sheet.Cells["A2"].PutValue(-456);  // Negative number
                sheet.Cells["A3"].PutValue(0);     // Zero (should appear as "-")
                sheet.Cells["A4"].PutValue(78.9);  // Positive decimal
                sheet.Cells["A5"].PutValue(0.0);   // Zero as decimal (should also appear as "-")

                // Save the workbook
                workbook.Save("CustomZeroFormat.xlsx");
                Console.WriteLine("Workbook saved successfully as CustomZeroFormat.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

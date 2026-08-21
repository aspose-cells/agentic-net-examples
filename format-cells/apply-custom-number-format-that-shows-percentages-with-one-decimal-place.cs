// Title: Apply a Custom Percentage Format (One Decimal) Using Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, writes the value 0.456 to cell A1, defines a style with the custom number format "0.0%" to display percentages with one decimal place, applies the style, and saves the file as PercentageOneDecimalDemo.xlsx.
// Keywords: Aspose.Cells | C# | .NET | custom number format | percentage format | one decimal place | cell styling | Excel export | formatting percentages
// Common Searches: Aspose.Cells format cell as percentage with one decimal | C# custom number format string for percentages | how to set percentage format 0.0% in Aspose.Cells | apply custom style to Excel cell using Aspose.Cells .NET
// Developer Intent: Format a worksheet cell to show a percentage value with a single decimal digit.
// Use Cases: Financial reports that require ratios displayed as 45.6% instead of 0.456. | Dashboard sheets where all percentage metrics must share a consistent one‑decimal format. | Automated Excel exports that preserve precise percentage representation for analytics.
// AI Prompts: Generate C# code with Aspose.Cells that applies the custom format "0.0%" to a range of cells. | Show how to change the format to two decimal places or add thousand separators in Aspose.Cells. | Explain the difference between built‑in percentage formats and custom formats in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This example creates a workbook, writes the value 0.456 to cell A1, defines a style with the custom number format "0.0%" to display percentages with one decimal place, applies the style, and saves the file as PercentageOneDecimalDemo.xlsx.
    public class PercentageOneDecimalDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Put a numeric value (e.g., 45.6%)
                Cell cell = sheet.Cells["A1"];
                cell.PutValue(0.456); // 45.6%

                // Create a style with a custom percentage format showing one decimal place
                Style percentStyle = workbook.CreateStyle();
                percentStyle.Custom = "0.0%"; // custom format: one decimal place percentage

                // Apply the style to the cell
                cell.SetStyle(percentStyle);

                // Determine output file path
                string outputFile = "PercentageOneDecimalDemo.xlsx";

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputFile));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            PercentageOneDecimalDemo.Run();
        }
    }
}

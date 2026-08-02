// Title: Aspose.Cells .NET: Custom number format to display negative percentages in red with a minus sign
// Description: Demonstrates how to create a workbook, insert positive and negative decimal values, define a custom number format "0.00%;[Red]-0.00%" that shows negative percentages in red with a leading minus sign, apply the format to a specific range using a StyleFlag, and save the file as an XLSX document.
// Keywords: Aspose.Cells custom number format | negative percentage red format .NET | C# Aspose.Cells style flag | Excel percentage formatting Aspose | display negative values in red | custom Excel number format string | Aspose.Cells tutorial | C# Excel styling example
// Common Searches: Aspose.Cells format negative percentages red | C# custom number format for percentages | How to apply style flag number format Aspose.Cells | Excel red negative values custom format .NET | Aspose.Cells example for percentage styling | Create workbook with colored negative percentages C#
// Developer Intent: Generate an Excel file where negative percentage cells appear in red with a minus sign using Aspose.Cells for .NET.
// Use Cases: Financial reports that highlight loss percentages in red for quick visual analysis. | Sales or KPI dashboards where decline rates need distinct coloring without extra conditional rules. | Automated data exports that require consistent visual formatting of positive and negative percentages.
// AI Prompts: Write C# code with Aspose.Cells to apply a custom number format that shows negative percentages in red and prefixed with a minus sign. | Show how to use StyleFlag to apply only the number format to a range while keeping other cell styles unchanged. | Explain how to modify the format string to use parentheses instead of a minus sign for negative percentages in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, insert positive and negative decimal values, define a custom number format "0.00%;[Red]-0.00%" that shows negative percentages in red with a leading minus sign, apply the format to a specific range using a StyleFlag, and save the file as an XLSX document.
    public class NegativePercentageFormatDemo
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put sample values (as decimal fractions)
            sheet.Cells["A1"].PutValue(0.25);   // 25%
            sheet.Cells["A2"].PutValue(-0.45); // -45%

            // Create a style with a custom number format:
            // Positive percentages normal, negative percentages red with a minus sign
            Style style = workbook.CreateStyle();
            style.Custom = "0.00%;[Red]-0.00%";

            // Apply only the number format part of the style
            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true;

            // Apply the style to the target range (A1:A2)
            Aspose.Cells.Range range = sheet.Cells.CreateRange("A1", "A2");
            range.ApplyStyle(style, flag);

            // Save the workbook
            string outputPath = "NegativePercentageFormatDemo.xlsx";
            workbook.Save(outputPath);
        }
    }
}

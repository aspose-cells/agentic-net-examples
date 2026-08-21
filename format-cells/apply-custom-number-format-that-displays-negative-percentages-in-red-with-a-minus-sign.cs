// Title: C# – Aspose.Cells: Custom number format to display negative percentages in red with a minus sign
// Description: Creates a workbook, inserts 25% and -45% values, defines a style with the custom format "0.00%;[Red]-0.00%" to show positive percentages normally and negative percentages in red with a leading minus sign, applies the style to cells A1:A2 using a StyleFlag that targets only the number format, and saves the file as an XLSX document.
// Keywords: Aspose.Cells custom number format | negative percentage red format | C# Excel style flag | Aspose.Cells .NET formatting | Excel red negative values | percentage display Aspose
// Common Searches: Aspose.Cells show negative percentages in red | custom number format string for red negative percentages C# | apply number format to a range Aspose.Cells | StyleFlag number format only Aspose.Cells | C# Excel red negative percent formatting
// Developer Intent: Apply a custom number format so that negative percentage values appear in red with a preceding minus sign.
// Use Cases: Financial statements where loss percentages are highlighted in red. | KPI dashboards that differentiate negative growth rates with colored percentages. | Automated Excel reports that embed visual cues for negative values without using conditional formatting.
// AI Prompts: Give a C# Aspose.Cells example that formats negative percentages in red with a minus sign using a custom number format. | Show how to modify the format string to display negative percentages in parentheses instead of a minus sign. | Demonstrate applying the red‑negative‑percentage format to an entire column programmatically.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    // Creates a workbook, inserts 25% and -45% values, defines a style with the custom format "0.00%;[Red]-0.00%" to show positive percentages normally and negative percentages in red with a leading minus sign, applies the style to cells A1:A2 using a StyleFlag that targets only the number format, and saves the file as an XLSX document.
    public class NegativePercentageNumberFormatDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
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

            // Set sample values (positive and negative percentages)
            sheet.Cells["A1"].PutValue(0.25);   // 25%
            sheet.Cells["A2"].PutValue(-0.45); // -45%

            // Create a style with a custom number format:
            // Positive percentages: 0.00%
            // Negative percentages: red color with minus sign
            Style style = workbook.CreateStyle();
            style.Custom = "0.00%;[Red]-0.00%";

            // Apply only the number format using a StyleFlag
            StyleFlag styleFlag = new StyleFlag();
            styleFlag.NumberFormat = true;

            // Apply the style to the target range (A1:A2)
            AsposeRange range = sheet.Cells.CreateRange("A1", "A2");
            range.ApplyStyle(style, styleFlag);

            // Save the workbook
            string outputPath = "NegativePercentageNumberFormatDemo.xlsx";
            workbook.Save(outputPath);
        }
    }
}

// Title: C# – Show a Plus Sign for Positive Percentages with a Custom Number Format in Aspose.Cells
// Description: Creates a workbook, writes values to A1‑A3, defines the custom format "+0.00%;-0.00%;0.00%" to display a leading ‘+’ for positive percentages, applies the format to the range using a StyleFlag that targets only the number format, and saves the file as an XLSX document.
// Keywords: Aspose.Cells custom format | C# percentage plus sign | Excel positive sign format | StyleFlag number format | .NET workbook formatting
// Common Searches: Aspose.Cells display + sign for positive percentages | custom number format string positive negative zero .NET | apply number format to a range Aspose.Cells C# | how to add plus sign to percentage values in Excel with code
// Developer Intent: Generate an Excel file where positive percentages are prefixed with ‘+’, while negative values keep ‘‑’ and zero values show no sign.
// Use Cases: Financial statements that highlight growth with a plus sign. | KPI dashboards where increases need explicit ‘+’ markers. | Automated reporting that distinguishes positive, negative, and zero percentages clearly.
// AI Prompts: Provide C# Aspose.Cells code that defines a custom number format "+0.00%;-0.00%;0.00%" and applies it to cells A1:A3. | Explain how to use StyleFlag to restrict formatting to the number format portion of a style in Aspose.Cells. | Show an example of saving a workbook after applying separate patterns for positive, negative, and zero percentages.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, writes values to A1‑A3, defines the custom format "+0.00%;-0.00%;0.00%" to display a leading ‘+’ for positive percentages, applies the format to the range using a StyleFlag that targets only the number format, and saves the file as an XLSX document.
    public class PercentageWithPlusSignDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Sample values: positive, negative, zero
                sheet.Cells["A1"].PutValue(0.25);   // 25%
                sheet.Cells["A2"].PutValue(-0.10); // -10%
                sheet.Cells["A3"].PutValue(0.0);   // 0%

                // Define a custom number format with plus sign for positives
                Style style = workbook.CreateStyle();
                style.Custom = "+0.00%;-0.00%;0.00%";

                // Apply the style to the range A1:A3
                Aspose.Cells.Range range = sheet.Cells.CreateRange("A1:A3");
                StyleFlag flag = new StyleFlag();
                flag.NumberFormat = true; // Apply only the number format part
                range.ApplyStyle(style, flag);

                // Save the workbook
                string outputPath = "PercentageWithPlusSignDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Program entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            PercentageWithPlusSignDemo.Run();
        }
    }
}

// Title: Apply Custom Scientific Notation (0.00E+00) Number Format to a Cell with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, insert a numeric value, define a style with the custom format "0.00E+00", use StyleFlag to modify only the number format, apply the style to cell A1, and save the Excel file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells scientific notation | C# custom number format | 0.00E+00 Excel format | StyleFlag NumberFormat | apply number format Aspose.Cells | Excel cell formatting .NET
// Common Searches: Aspose.Cells set cell to scientific notation C# | custom number format 0.00E+00 Aspose.Cells | how to use StyleFlag to change only number format | format large numbers as 1.23E+04 in .NET Excel | apply number format without affecting other styles
// Developer Intent: The developer needs to display a numeric value in scientific notation with two decimal places while leaving other cell styling unchanged.
// Use Cases: Present engineering measurements compactly (e.g., 1.23E+04) in reports. | Show scientific data with consistent two‑decimal precision across selected cells. | Export financial calculations where exponential notation improves readability without altering fonts or borders.
// AI Prompts: Write C# code using Aspose.Cells to apply the "0.00E+00" format to a specific range while preserving existing cell styles. | Explain the role of StyleFlag.NumberFormat in Aspose.Cells and give a concise example. | Provide a step‑by‑step tutorial for creating a workbook, inserting a value, setting a scientific notation style, and saving the file with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, insert a numeric value, define a style with the custom format "0.00E+00", use StyleFlag to modify only the number format, apply the style to cell A1, and save the Excel file using Aspose.Cells for .NET.
    public class ScientificNotationNumberFormatDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Set a numeric value in cell A1
                sheet.Cells["A1"].PutValue(12345.6789);

                // Create a style with a custom scientific notation format (two decimal places)
                Style style = workbook.CreateStyle();
                style.Custom = "0.00E+00";

                // Apply only the number format part of the style
                StyleFlag styleFlag = new StyleFlag();
                styleFlag.NumberFormat = true;

                // Apply the style to cell A1
                AsposeRange range = sheet.Cells.CreateRange("A1");
                range.ApplyStyle(style, styleFlag);

                // Ensure the output directory exists
                string outputPath = "ScientificNotationNumberFormatDemo.xlsx";
                string directory = System.IO.Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(directory) && !System.IO.Directory.Exists(directory))
                {
                    System.IO.Directory.CreateDirectory(directory);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}

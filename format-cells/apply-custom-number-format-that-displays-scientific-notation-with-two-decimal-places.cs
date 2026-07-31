// Title: Apply scientific notation (two decimal places) with a custom number format in Aspose.Cells for .NET (C#)
// Description: Shows how to create a workbook, write a numeric value, define a style with the custom format "0.00E+00", use a StyleFlag to apply only the number‑format, and save the file as ScientificFormat.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | custom number format | scientific notation | two decimal places | StyleFlag | Excel export | format cell | 0.00E+00
// Common Searches: Aspose.Cells format cell as scientific notation | C# set custom number format Excel Aspose | apply number format only with StyleFlag Aspose.Cells | custom scientific notation 0.00E+00 Aspose | how to display numbers in scientific notation using Aspose.Cells
// Developer Intent: Create a workbook and format a specific cell to display numbers in scientific notation with two decimal places.
// Use Cases: Financial models that require large numbers to be shown consistently in scientific notation. | Engineering or scientific reports exporting measurement data to Excel with a fixed scientific format. | Data‑logging applications where downstream systems expect values in scientific notation. | Spreadsheet templates that automatically enforce scientific notation on user‑entered values.
// AI Prompts: Provide C# code to apply the custom format "0.000E+00" to an entire column using Aspose.Cells. | Show how to combine a scientific notation format with bold font and a yellow background in Aspose.Cells. | Explain how to change the number‑format string at runtime based on a user‑selected option in an Aspose.Cells workbook. | Generate a PowerShell script that uses Aspose.Cells to apply a scientific notation format to a range of cells.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, write a numeric value, define a style with the custom format "0.00E+00", use a StyleFlag to apply only the number‑format, and save the file as ScientificFormat.xlsx using Aspose.Cells for .NET.
    public class ApplyScientificNumberFormat
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Put a numeric value into cell A1
                sheet.Cells["A1"].PutValue(12345.6789);

                // Create a style with a custom scientific notation format (two decimal places)
                Style style = workbook.CreateStyle();
                style.Custom = "0.00E+00";

                // Configure a StyleFlag to apply only the number format part of the style
                StyleFlag styleFlag = new StyleFlag();
                styleFlag.NumberFormat = true;

                // Apply the style to cell A1 using the defined StyleFlag
                AsposeRange range = sheet.Cells.CreateRange("A1");
                range.ApplyStyle(style, styleFlag);

                // Save the workbook to a file
                string outputPath = "ScientificFormat.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            ApplyScientificNumberFormat.Run();
        }
    }
}

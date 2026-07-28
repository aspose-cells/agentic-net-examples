// Title: C# – Apply custom scientific notation (three significant digits) with Aspose.Cells
// Description: Demonstrates how to create a workbook, insert a large number, define a style with the custom format "0.00E+00" (three significant digits), apply it to a cell or range, and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | custom number format | scientific notation | three significant digits | 0.00E+00 | cell style | Excel export | formatting large numbers
// Common Searches: Aspose.Cells scientific notation three digits | C# custom number format 0.00E+00 | apply number format to cell Aspose.Cells | how to show three significant digits in Excel with Aspose | set cell style format in Aspose.Cells .NET
// Developer Intent: The developer needs to format one or more cells so that numeric values appear in scientific notation with exactly three significant digits.
// Use Cases: Present large financial figures in a compact scientific format. | Standardize engineering measurement columns with three‑digit scientific notation. | Export scientific experiment results to Excel where concise notation is required.
// AI Prompts: Generate C# code that applies the custom format "0.00E+00" to a specified range using Aspose.Cells. | Explain how to modify the format string to show four significant digits instead of three. | Show how to combine the scientific notation format with additional styling such as font color, bold text, and borders.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, insert a large number, define a style with the custom format "0.00E+00" (three significant digits), apply it to a cell or range, and save the file using Aspose.Cells for .NET.
    class ScientificNotationThreeSigDigits
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Put a numeric value that will be displayed in scientific notation
            Cell cell = sheet.Cells["A1"];
            cell.PutValue(123456789.0);

            // Create a style with a custom scientific format showing three significant digits
            // Format: one digit before decimal, two after (total three significant digits)
            Style style = workbook.CreateStyle();
            style.Custom = "0.00E+00"; // e.g., 1.23E+08

            // Apply only the number format part of the style
            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true;

            // Apply the style to the target cell
            cell.SetStyle(style);
            // Alternatively, for a range:
            // Aspose.Cells.Range range = sheet.Cells.CreateRange("A1");
            // range.ApplyStyle(style, flag);

            // Save the workbook (lifecycle save rule)
            workbook.Save("ScientificNotationThreeSigDigits.xlsx");

            // Optional: output confirmation
            Console.WriteLine("Workbook saved with scientific notation format applied to A1.");
        }
    }
}

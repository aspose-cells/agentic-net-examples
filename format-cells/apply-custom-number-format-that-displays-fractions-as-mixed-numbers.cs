// Title: Apply Mixed Fraction Custom Number Format (# ?/?) with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, writes 2.75 to cell A1, defines a style with the custom format "# ?/?" to display values as mixed fractions, applies the style using a StyleFlag that targets only the number format, saves the file as MixedFractionFormatDemo.xlsx, and prints the formatted string to the console.
// Keywords: Aspose.Cells | mixed fraction format | custom number format | # ?/? | C# | .NET | Excel cell formatting | fraction display
// Common Searches: Aspose.Cells mixed fraction custom format | C# apply "# ?/?" number format in Excel | how to show fractions as mixed numbers with Aspose.Cells | format cell as mixed fraction .NET | custom number format for fractions Aspose
// Developer Intent: Display numeric values as mixed fractions by applying a custom number format to one or more worksheet cells.
// Use Cases: Financial statements that require measurements like 2 3/4 units for clarity. | Engineering reports where dimensions are traditionally expressed as mixed fractions. | Invoices or order forms that list quantities in mixed‑fraction notation.
// AI Prompts: Generate C# code using Aspose.Cells to apply the "# ?/?" mixed‑fraction format to an entire column while preserving existing cell styles. | Explain how to modify the custom format to show leading zeros in the fractional part (e.g., "# 01/04"). | Provide a sample that formats multiple rows with the mixed‑fraction style and then exports the workbook to PDF.

using System;
using Aspose.Cells;

namespace AsposeCellsFractionMixedNumberDemo
{
    // Creates a workbook, writes 2.75 to cell A1, defines a style with the custom format "# ?/?" to display values as mixed fractions, applies the style using a StyleFlag that targets only the number format, saves the file as MixedFractionFormatDemo.xlsx, and prints the formatted string to the console.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set a numeric value that will be displayed as a mixed fraction
            // Example: 2.75 will be shown as "2 3/4"
            sheet.Cells["A1"].PutValue(2.75);

            // Create a style with a custom number format for mixed fractions
            // "# ?/?" displays the integer part followed by a simple fraction
            Style mixedFractionStyle = workbook.CreateStyle();
            mixedFractionStyle.Custom = "# ?/?";

            // Use StyleFlag to apply only the number format part of the style
            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true;

            // Apply the style to the target cell (A1)
            // CreateRange(row, column, totalRows, totalColumns)
            Aspose.Cells.Range targetRange = sheet.Cells.CreateRange(0, 0, 1, 1);
            targetRange.ApplyStyle(mixedFractionStyle, flag);

            // Save the workbook to a file
            workbook.Save("MixedFractionFormatDemo.xlsx");

            // Optional: Output the formatted string to console for verification
            Console.WriteLine("Formatted value in A1: " + sheet.Cells["A1"].StringValue);
        }
    }
}

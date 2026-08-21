// Title: Apply custom number format '#,##0;[Red]-#,##0' to a cell range using Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, defines a style with the custom format "#,##0;[Red]-#,##0" to display positive numbers normally and negative numbers in red with a minus sign, applies the style only to the number format via a StyleFlag, formats the range A1:A5, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells | C# | .NET | custom number format | red negative numbers | StyleFlag | ApplyStyle | Excel formatting | pivot table source formatting | financial report styling
// Common Searches: Aspose.Cells set red color for negative numbers C# | custom number format '#,##0;[Red]-#,##0' Aspose.Cells | apply number format to a range Aspose.Cells .NET | use StyleFlag to change only number format in Aspose.Cells | format negative values in red with Aspose.Cells
// Developer Intent: Format a cell range so that negative values appear in red while leaving other style attributes unchanged.
// Use Cases: Highlight losses in financial statements by showing negative amounts in red. | Standardize numeric display across exported Excel reports. | Prepare source data for a pivot table with consistent number formatting. | Create visually distinct dashboards where negative metrics are emphasized.
// AI Prompts: Generate C# code that creates a workbook, defines a style with the custom number format "#,##0;[Red]-#,##0", and applies it to range A1:A5 using StyleFlag in Aspose.Cells. | Show how to change only the number format of cells A1:A10 to red‑negative style while preserving existing fonts, borders, and fill colors. | Explain how to modify the custom format to use parentheses for negative numbers instead of a minus sign in Aspose.Cells.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsNumberFormatDemo
{
    // This example creates a workbook, defines a style with the custom format "#,##0;[Red]-#,##0" to display positive numbers normally and negative numbers in red with a minus sign, applies the style only to the number format via a StyleFlag, formats the range A1:A5, and saves the file as an Excel workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample numeric data (positive, negative and zero)
                sheet.Cells["A1"].PutValue(12345);
                sheet.Cells["A2"].PutValue(-6789);
                sheet.Cells["A3"].PutValue(0);
                sheet.Cells["A4"].PutValue(2500);
                sheet.Cells["A5"].PutValue(-1500);

                // Define a custom number format:
                // Positive numbers: 1,2345
                // Negative numbers: red with a minus sign
                Style customStyle = workbook.CreateStyle();
                customStyle.Custom = "#,##0;[Red]-#,##0";

                // Apply only the number format, preserving other style attributes
                StyleFlag flag = new StyleFlag { NumberFormat = true };

                // Apply the style to the target range
                AsposeRange range = sheet.Cells.CreateRange("A1:A5");
                range.ApplyStyle(customStyle, flag);

                // Save the workbook
                workbook.Save("CustomNumberFormatDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}

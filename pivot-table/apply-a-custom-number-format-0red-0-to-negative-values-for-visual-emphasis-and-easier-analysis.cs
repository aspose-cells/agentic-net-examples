// Title: C# – Apply '#,##0;[Red]-#,##0' custom number format to highlight negative values with Aspose.Cells
// Description: Demonstrates how to create a workbook, insert positive and negative numbers, define a style using the custom format "#,##0;[Red]-#,##0", apply only the number‑format part with StyleFlag, and save the file as CustomNumberFormat.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells custom number format | C# Excel negative numbers red | StyleFlag number format only | apply custom format Aspose.Cells | highlight negative values Excel | Aspose.Cells .NET tutorial | Excel formatting with Aspose | red negative numbers C# | range style application Aspose | custom numeric display Excel
// Common Searches: Aspose.Cells set custom number format for negatives | C# apply red negative number style in Excel | How to use StyleFlag to change only number format | Excel custom format '#,##0;[Red]-#,##0' Aspose | Apply number format to a cell range with Aspose.Cells
// Developer Intent: Apply a red‑negative custom number format to a specific cell range in an Excel workbook.
// Use Cases: Show financial losses in red to improve report readability. | Standardize numeric appearance across multiple worksheets. | Make negative inventory or KPI values instantly recognizable.
// AI Prompts: Generate C# code that applies the '#,##0;[Red]-#,##0' format to column B rows 2‑10 using Aspose.Cells. | Explain how StyleFlag can isolate number‑format changes without altering fonts or borders in Aspose.Cells. | Create a reusable method that accepts a worksheet, range address, and applies a red‑negative number format.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, insert positive and negative numbers, define a style using the custom format "#,##0;[Red]-#,##0", apply only the number‑format part with StyleFlag, and save the file as CustomNumberFormat.xlsx using Aspose.Cells for .NET.
class ApplyCustomNumberFormat
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data with positive and negative numbers
            sheet.Cells["A1"].PutValue(1234);
            sheet.Cells["A2"].PutValue(-5678);
            sheet.Cells["A3"].PutValue(91011);
            sheet.Cells["A4"].PutValue(-1213);
            sheet.Cells["A5"].PutValue(1415);

            // Create a style with the custom number format for negative values
            Style customStyle = workbook.CreateStyle();
            customStyle.Custom = "#,##0;[Red]-#,##0";

            // Use StyleFlag to apply only the number format part of the style
            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true;

            // Apply the style to the range containing the data
            Aspose.Cells.Range range = sheet.Cells.CreateRange("A1", "A5");
            range.ApplyStyle(customStyle, flag);

            // Save the workbook
            workbook.Save("CustomNumberFormat.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

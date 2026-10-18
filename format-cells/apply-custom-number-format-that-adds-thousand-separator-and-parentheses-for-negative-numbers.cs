// Title: Apply a custom numeric format with thousand separators and parentheses for negative numbers using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set custom number format with thousand separator and parentheses for negatives | How to display negative values in parentheses in an Excel file using Aspose.Cells .NET | C# code to apply "#,##0;(#,##0)" format to cells with Aspose.Cells | Formatting numbers with commas and parentheses for negatives in Aspose.Cells workbook | Create and apply custom style for numeric formatting in Aspose.Cells for .NET
// Tags: custom numeric format Aspose.Cells | thousand separator formatting C# | negative numbers parentheses Aspose.Cells | cell style custom number format .NET | Excel custom number format Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a workbook, writes a positive and a negative integer to cells A1 and A2, defines the custom format "#,##0;(#,##0)" to add thousand separators and show negatives in parentheses, applies this style to both cells, and saves the file as CustomNumberFormat.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Put a positive and a negative number in cells A1 and A2
        sheet.Cells["A1"].PutValue(1234567);
        sheet.Cells["A2"].PutValue(-1234567);

        // Define a custom number format:
        //   - "#,##0" adds thousand separators for positive numbers
        //   - "(#,##0)" displays negative numbers in parentheses
        string customFormat = "#,##0;(#,##0)";

        // Apply the custom format to cell A1
        Style stylePos = workbook.CreateStyle();
        stylePos.Custom = customFormat;
        sheet.Cells["A1"].SetStyle(stylePos);

        // Apply the same custom format to cell A2
        Style styleNeg = workbook.CreateStyle();
        styleNeg.Custom = customFormat;
        sheet.Cells["A2"].SetStyle(styleNeg);

        // Save the workbook to a file
        workbook.Save("CustomNumberFormat.xlsx");
    }
}

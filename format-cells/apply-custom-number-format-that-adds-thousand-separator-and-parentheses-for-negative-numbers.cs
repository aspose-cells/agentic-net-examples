// Title: How to apply a custom number format with thousand separators and parentheses for negative values using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set custom number format #,##0;(#,##0) for Excel cells | C# apply thousand separator and parentheses to negative numbers in Excel with Aspose.Cells | Create a style with a custom numeric pattern in Aspose.Cells .NET example | Formatting positive and negative values using a custom style in an Aspose.Cells workbook | Save an Excel file with a custom number format using Aspose.Cells for .NET
// Tags: custom numeric format Aspose.Cells .NET | thousand separator format Aspose.Cells | negative number parentheses format Excel C# | style custom number pattern Aspose.Cells | apply custom style to worksheet cells Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a workbook, writes a positive and a negative number to cells A1 and A2, defines a custom style with the pattern "#,##0;(#,##0)" to add thousand separators and wrap negative values in parentheses, applies the style to both cells, and saves the file as CustomNumberFormat.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Insert sample positive and negative numbers
        sheet.Cells["A1"].PutValue(12345.67);
        sheet.Cells["A2"].PutValue(-12345.67);

        // Create a style and set a custom number format:
        // "#,##0" for positive numbers (adds thousand separator)
        // "(#,##0)" for negative numbers (adds parentheses)
        Style customStyle = workbook.CreateStyle();
        customStyle.Custom = "#,##0;(#,##0)";

        // Apply the custom style to the cells
        sheet.Cells["A1"].SetStyle(customStyle);
        sheet.Cells["A2"].SetStyle(customStyle);

        // Save the workbook to a file
        workbook.Save("CustomNumberFormat.xlsx");
    }
}

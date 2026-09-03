// Title: Apply a custom number format in Aspose.Cells for .NET to show negative values in red parentheses
// AI Prompts: Write C# code using Aspose.Cells that creates a workbook, inserts a negative number, and applies a custom numeric format where negatives appear in red parentheses while positives keep the default format. | Show how to define a Style with a custom format string '#,##0.00_);[Red](#,##0.00)' and apply it to a cell using a StyleFlag that only changes the number format in Aspose.Cells.
// Common Searches: how to display negative numbers in red parentheses using Aspose.Cells for .NET | Aspose.Cells C# custom numeric format for negative values with red color and parentheses | set Excel cell style to show negatives in red parentheses with Aspose.Cells | define custom number format '#,##0.00_);[Red](#,##0.00)' in Aspose.Cells workbook | apply style flag to change only number format in Aspose.Cells C#
// Tags: custom numeric format red parentheses Aspose.Cells | negative value formatting Excel .NET | style flag number format Aspose.Cells C# | apply custom number format workbook Aspose.Cells | excel cell style red parentheses negative numbers

using System;
using Aspose.Cells;

// Creates a workbook, puts a negative value in A1, defines a custom number format '#,##0.00_);[Red](#,##0.00)' to render negatives in red parentheses, applies the style with a StyleFlag limited to number formatting, and saves the file as CustomNumberFormat.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Put a negative value into cell A1
        Cell cell = sheet.Cells["A1"];
        cell.PutValue(-1234.56);

        // Create a style object
        Style style = workbook.CreateStyle();

        // Set custom number format:
        // Positive numbers: normal format
        // Negative numbers: red color, enclosed in parentheses
        // Zero: display as 0.00
        style.Custom = "#,##0.00_);[Red](#,##0.00)";

        // Apply the style to the cell
        StyleFlag flag = new StyleFlag();
        flag.NumberFormat = true; // Apply only number format
        cell.SetStyle(style, flag);

        // Save the workbook to a file
        workbook.Save("CustomNumberFormat.xlsx");
    }
}

// Title: Apply a custom number format with thousand separators and parentheses for negative values using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a style in Aspose.Cells using the custom format "#,##0;(#,##0)" and applies it to specific cells to display commas and parentheses for negatives. | Write a snippet that sets positive and negative numeric values in a worksheet, assigns the custom numeric style, and saves the workbook as an .xlsx file.
// Common Searches: Aspose.Cells C# custom numeric format with commas and parentheses for negative numbers | How to display negative numbers in parentheses using Aspose.Cells .NET | C# apply thousand separator format to cells in Aspose.Cells workbook | Define and use custom number style "#,##0;(#,##0)" in Aspose.Cells | Saving a workbook after applying custom number format in Aspose.Cells C#
// Tags: custom numeric format commas Aspose.Cells | parentheses negative number style C# | cell style for thousand separators Aspose.Cells | apply custom format to worksheet cells .NET | export workbook with styled numbers Aspose.Cells

using System;
using Aspose.Cells;

// Demonstrates creating a workbook, inserting positive and negative values, defining a custom number format "#,##0;(#,##0)" that adds thousand separators and encloses negatives in parentheses, applying the style to the cells, and saving the file as CustomNumberFormat.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set a positive and a negative value for demonstration
        Cell positiveCell = sheet.Cells["A1"];
        positiveCell.PutValue(1234567.89);

        Cell negativeCell = sheet.Cells["A2"];
        negativeCell.PutValue(-1234567.89);

        // Define a custom number format:
        // "#,##0;(#,##0)" adds thousand separators and encloses negative numbers in parentheses
        string customFormat = "#,##0;(#,##0)";

        // Apply the custom format to both cells
        Style style = workbook.CreateStyle();
        style.Custom = customFormat;

        // Assign the style to the cells
        positiveCell.SetStyle(style);
        negativeCell.SetStyle(style);

        // Save the workbook to a file
        workbook.Save("CustomNumberFormat.xlsx");
    }
}

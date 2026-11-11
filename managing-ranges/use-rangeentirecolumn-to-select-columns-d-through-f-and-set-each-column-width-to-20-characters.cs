// Title: Set column width to 20 characters for columns D‑F using Range.EntireColumn in Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set width of columns D to F programmatically | How to use Range.EntireColumn to adjust multiple column widths in a .NET workbook | Aspose.Cells example for setting column width to a specific character count | Apply same column width to several columns with Aspose.Cells C# API | Set Excel column width to 20 characters for D, E, and F using Aspose.Cells range
// Tags: range.entirecolumn set column width c# | aspocells column width adjustment | c# set multiple excel columns width | excel column width via aspocells range | set column width 20 characters aspocells | c# aspocells entirecolumn usage

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Creates a workbook, defines a range covering D1:F1, uses the EntireColumn property to select columns D, E, and F, sets each column's width to 20 characters, and saves the file as Result.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Create a range that covers cells D1 to F1
            AsposeRange range = sheet.Cells.CreateRange("D1:F1");

            // Use EntireColumn to select the whole columns D, E, and F
            // and set each column width to 20 characters
            range.EntireColumn.ColumnWidth = 20;

            // Save the workbook
            workbook.Save("Result.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

// Title: Auto‑fit a single column in an Aspose.Cells worksheet based on its cell contents using C#
// AI Prompts: Generate C# code that creates a workbook, fills column A with strings of varying length, prints the column width before and after calling Worksheet.AutoFitColumn for rows 0‑2, and saves the file. | Write a reusable C# method that accepts a Worksheet, column index, start row, end row, and optionally logs the column width before and after auto‑fitting with Aspose.Cells. | Show how to use Worksheet.AutoFitColumn to adjust the width of column 0 for a specific row range and then retrieve the updated column width in C#.
// Common Searches: C# Aspose.Cells how to auto fit a column for a specific row range | retrieve column width before and after AutoFitColumn in Aspose.Cells | example of using Worksheet.AutoFitColumn with start and end row indices | auto fit column based on cell content Aspose.Cells .NET example
// Tags: Aspose.Cells Worksheet.AutoFitColumn for specific rows | auto-fit column width based on cell content C# | retrieve column width Aspose.Cells | save workbook after column auto-fit Aspose.Cells | C# example auto-fitting single column Aspose.Cells

using System;
using Aspose.Cells;

// The sample creates a new workbook, populates column A with short and long text strings, displays the column width before and after invoking worksheet.AutoFitColumn(0, 0, 2) to fit rows 0‑2, and saves the result as AutoFitColumnDemo.xlsx.
class AutoFitColumnExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate column A (index 0) with sample data of varying lengths
        worksheet.Cells["A1"].PutValue("Short");
        worksheet.Cells["A2"].PutValue("This is a longer text that should cause the column to expand");
        worksheet.Cells["A3"].PutValue("Another very very long text entry to demonstrate auto-fitting of column width");

        // Show column width before auto‑fit (optional)
        Console.WriteLine($"Column A width before AutoFitColumn: {worksheet.Cells.GetColumnWidth(0)}");

        // Auto‑fit column A for rows 0 to 2 (zero‑based indices)
        worksheet.AutoFitColumn(0, 0, 2);

        // Show column width after auto‑fit (optional)
        Console.WriteLine($"Column A width after AutoFitColumn: {worksheet.Cells.GetColumnWidth(0)}");

        // Save the workbook to a file
        workbook.Save("AutoFitColumnDemo.xlsx");
    }
}

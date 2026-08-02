// Title: Merge Top Row Cells and Freeze Header Row with Aspose.Cells for .NET (C#)
// Description: This example demonstrates how to create a new workbook, merge cells A1:E1 into a single header, set its text, freeze the first row so the merged header stays visible while scrolling, and save the file as MergedHeaderFreeze.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | merge cells | freeze panes | header row | Excel worksheet | merged header | freeze top row | Workbook.Save
// Common Searches: Aspose.Cells merge top row C# | freeze header row after merging cells Aspose.Cells | how to freeze panes with merged header in Excel using .NET | C# Aspose.Cells merge A1:E1 and freeze first row | create frozen merged header in Excel with Aspose.Cells
// Developer Intent: Create an Excel workbook, merge the first row across multiple columns, and freeze that header row to keep it visible during scrolling.
// Use Cases: Design a report where the title spans several columns and remains fixed at the top while users scroll through data. | Build a reusable worksheet template with a merged, frozen header to improve navigation in large datasets. | Develop a dashboard sheet where a merged header provides context and stays visible during both vertical and horizontal scrolling.
// AI Prompts: Generate C# code with Aspose.Cells that merges cells A1 through E1, sets a header value, and freezes the first row. | Show an Aspose.Cells example for merging the top row across five columns and applying FreezePanes to keep the header visible. | Explain the parameters of Worksheet.FreezePanes when freezing a merged header row in an Excel file using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This example demonstrates how to create a new workbook, merge cells A1:E1 into a single header, set its text, freeze the first row so the merged header stays visible while scrolling, and save the file as MergedHeaderFreeze.xlsx using Aspose.Cells for .NET.
    class MergeAndFreezeHeader
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Merge the top row across five columns (A1:E1)
                cells.Merge(0, 0, 1, 5);
                cells[0, 0].PutValue("Header");

                // Freeze the top row so the merged header stays visible while scrolling
                // Freeze at row index 1 (second row), column index 0, freezing 1 row and 0 columns
                worksheet.FreezePanes(1, 0, 1, 0);

                // Save the workbook
                workbook.Save("MergedHeaderFreeze.xlsx");
                Console.WriteLine("Workbook saved successfully as MergedHeaderFreeze.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            MergeAndFreezeHeader.Run();
        }
    }
}

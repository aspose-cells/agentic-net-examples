// Title: Apply an AutoFilter to a header row and freeze that row with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a worksheet, adds a header row, applies an AutoFilter to the header, and then freezes the first row so the filter dropdowns stay visible using Aspose.Cells. | Update an existing Aspose.Cells workbook to set a dynamic AutoFilter range based on the populated data and keep the header row frozen when the file is opened.
// Common Searches: asp.net cells c# apply autofilter to first row and freeze pane | how to keep Excel filter dropdowns visible with Aspose.Cells | set autofilter range and freeze top row programmatically Aspose.Cells | freeze panes after applying autofilter using Aspose.Cells C# | Aspose.Cells keep filter controls accessible by freezing header row
// Tags: Aspose.Cells AutoFilter header row | Aspose.Cells freeze panes first row | C# set AutoFilter range Excel | C# freeze top row Aspose.Cells | Aspose.Cells preserve filter controls

using System;
using Aspose.Cells;

// The example creates a new workbook, inserts sample data with a header row, applies an AutoFilter to the range A1:B3, freezes the first row so the filter controls remain visible, and saves the file as AutoFilterAndFreeze.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Get the first worksheet
            var sheet = workbook.Worksheets[0];

            // Sample data with a header row in the first row
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Jane");
            sheet.Cells["B3"].PutValue(25);

            // Apply AutoFilter to the range that includes the header row
            sheet.AutoFilter.Range = "A1:B3";

            // Freeze the header row (first row) so filter controls stay visible
            // FreezePanes(row, column, totalRows, totalColumns)
            // Row = 1 (freeze rows above the second row), Column = 0 (no column freeze)
            sheet.FreezePanes(1, 0, 0, 0);

            // Save the workbook
            workbook.Save("AutoFilterAndFreeze.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

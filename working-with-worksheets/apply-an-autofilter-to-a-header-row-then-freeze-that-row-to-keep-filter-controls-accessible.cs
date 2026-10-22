// Title: How to apply an AutoFilter to a header row and freeze that row using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a new workbook, adds a header row, defines an AutoFilter over the header range, freezes the header row with FreezePanes, and saves the file using Aspose.Cells. | Show the exact Aspose.Cells API calls required to set an AutoFilter range and to freeze the first worksheet row while keeping filter controls visible.
// Common Searches: Aspose.Cells C# set AutoFilter on first row and keep filter dropdowns visible | C# freeze top row after applying AutoFilter with Aspose.Cells | How to use FreezePanes together with AutoFilter range in Aspose.Cells for .NET | Apply AutoFilter to header and freeze pane in generated Excel using Aspose.Cells C#
// Tags: autofilter range definition Aspose.Cells C# | freeze top row panes Aspose.Cells C# | apply autofilter and freeze header Aspose.Cells | Excel .xlsx generation with AutoFilter Aspose.Cells | worksheet FreezePanes usage Aspose.Cells C#

using Aspose.Cells;
using System;

// The example creates a new workbook, writes a header and sample data, sets an AutoFilter over the range A1:C3, freezes the first row with FreezePanes to keep filter controls visible, and saves the file as AutoFilterAndFreeze.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate header row
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Score");

            // Add sample data
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["C2"].PutValue(85);
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");
            sheet.Cells["C3"].PutValue(92);
            // Additional rows can be added similarly

            // Apply AutoFilter to the header row (first row) covering all data rows
            sheet.AutoFilter.Range = "A1:C3";

            // Freeze the header row so filter controls stay visible
            // Freeze rows above row index 1 (i.e., the first row) and no columns
            // Use the overload that includes total rows and columns to avoid API mismatch
            int totalRows = sheet.Cells.MaxDataRow + 1;      // total rows in the sheet
            int totalColumns = sheet.Cells.MaxDataColumn + 1; // total columns in the sheet
            sheet.FreezePanes(1, 0, totalRows, totalColumns);

            // Save the workbook
            workbook.Save("AutoFilterAndFreeze.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

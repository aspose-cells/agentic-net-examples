// Title: How to display filter arrows on a worksheet header row using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a workbook, adds a header row, sets the AutoFilter range to that header, calls Refresh, and saves the file so the filter arrows are visible. | Write a C# snippet that applies an AutoFilter to the first row of an Excel worksheet with Aspose.Cells and activates the sorting dropdowns programmatically.
// Common Searches: asp.net cells c# enable filter arrows on excel header row | set autofilter range and refresh aspose.cells c# example | programmatically show excel filter dropdowns using Aspose.Cells in C#
// Tags: Aspose.Cells AutoFilter activation C# | C# set AutoFilter range Excel worksheet | Aspose.Cells display Excel filter UI | Excel header sorting dropdown Aspose.Cells | programmatic Excel filter arrows C#

using Aspose.Cells;

// Creates a new workbook, adds a header row, defines the AutoFilter range covering the header cells, refreshes the filter to show filter arrows, and saves the workbook as FilteredWorkbook.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data with a header row
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Age");
        sheet.Cells["A2"].PutValue("John");
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["A3"].PutValue("Alice");
        sheet.Cells["B3"].PutValue(25);

        // Define the auto‑filter range covering the header row (A1:B1)
        sheet.AutoFilter.Range = "A1:B1";

        // Refresh the filter to display filter arrows on the header cells
        sheet.AutoFilter.Refresh();

        // Save the workbook to a file
        workbook.Save("FilteredWorkbook.xlsx");
    }
}

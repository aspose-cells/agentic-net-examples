// Title: Set landscape page orientation and freeze the header row in an Excel worksheet with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set worksheet to landscape orientation and keep header row visible | How to freeze the first row after changing page orientation with Aspose.Cells for .NET | C# example for setting page orientation to landscape and freezing panes in Excel using Aspose.Cells | Freeze panes on header row when using custom page layout in Aspose.Cells | Aspose.Cells .NET code to apply landscape page setup and freeze top row
// Tags: landscape orientation worksheet Aspose.Cells | freeze header row panes Aspose.Cells C# | page setup orientation Aspose.Cells .NET | freeze panes first row Excel Aspose.Cells | custom page layout Excel Aspose.Cells

using Aspose.Cells;
using System;

// Creates a new workbook, sets the first worksheet's page orientation to landscape, freezes the top header row so it remains visible while scrolling, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set a custom page orientation (Landscape)
        sheet.PageSetup.Orientation = PageOrientationType.Landscape;

        // Freeze the first row (header) so it stays visible when scrolling
        // Parameters: (row, column, totalRows, totalColumns)
        // Freeze rows above row index 1 (i.e., first row)
        sheet.FreezePanes(1, 0, 1, 0);

        // Save the workbook to a file
        workbook.Save("Output.xlsx", SaveFormat.Xlsx);
    }
}

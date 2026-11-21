// Title: Set worksheet zoom to 75% with Aspose.Cells for .NET (C# example)
// Common Searches: Aspose.Cells C# set worksheet zoom to 75 percent programmatically | How to change Excel worksheet view scale using Aspose.Cells .NET | C# code to adjust worksheet zoom level in generated Excel file | Set workbook worksheet zoom factor with Aspose.Cells API | Increase visible area of Excel sheet by setting zoom in Aspose.Cells
// Tags: worksheet zoom property Aspose.Cells | adjust worksheet view scale .NET | set Excel zoom level programmatically C# | Aspose.Cells workbook view settings | zoom factor configuration Aspose.Cells

using Aspose.Cells;
using System;

// Creates a new workbook, accesses the first worksheet, sets its Zoom property to 75, and saves the file as ZoomedWorksheet.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set the zoom factor to 75%
        sheet.Zoom = 75;

        // Save the workbook to a file
        workbook.Save("ZoomedWorksheet.xlsx");
    }
}

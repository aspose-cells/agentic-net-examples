// Title: Set a specific worksheet column width in pixels using Aspose.Cells Cells.SetColumnWidthPixel in C#
// Common Searches: C# Aspose.Cells set column B width to 100 pixels | How to use Cells.SetColumnWidthPixel to define column width in an Excel file with .NET | Set Excel column width by pixel value using Aspose.Cells API | Example of adjusting worksheet column width in pixels with Aspose.Cells for C# | Aspose.Cells Cells.SetColumnWidthPixel method usage for precise column sizing | Change column width to exact pixel count in a generated workbook using Aspose.Cells
// Tags: cells.setcolumnwidthpixel c# example | excel column width pixel specification aspnet | adjust worksheet column width pixels | define column width in pixels aspose.cells | set specific column width aspose.cells .net | precise column sizing pixel aspose

using Aspose.Cells;

// Creates a new workbook, accesses the first worksheet, sets column B (index 1) to a width of 100 pixels with Cells.SetColumnWidthPixel, and saves the workbook as Result.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set the width of column B (index 1) to 100 pixels
        sheet.Cells.SetColumnWidthPixel(1, 100);

        // Save the workbook
        workbook.Save("Result.xlsx");
    }
}

// Title: How to set an Excel column width in pixels with Aspose.Cells C# using Cells.SetColumnWidthPixel
// Common Searches: Aspose.Cells C# set column width to a specific pixel value | Set column B width to 120 pixels in an Excel workbook using Aspose.Cells | C# example of Cells.SetColumnWidthPixel method in Aspose.Cells | Adjust Excel column size by pixels with Aspose.Cells library | Programmatically define column width in pixels in a .NET workbook
// Tags: Aspose.Cells set column width pixel | C# Excel column width pixels | Cells.SetColumnWidthPixel method | adjust column size Aspose.Cells workbook | define column width in pixels .NET

using Aspose.Cells;

// This C# example creates a new Workbook, accesses the first worksheet, and uses Cells.SetColumnWidthPixel to set column B (index 1) to a width of 120 pixels, then saves the file as ColumnWidthExample.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // create

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set the width of column B (index 1) to 120 pixels
        sheet.Cells.SetColumnWidthPixel(1, 120); // set column width in pixels

        // Save the workbook to a file
        workbook.Save("ColumnWidthExample.xlsx"); // save
    }
}

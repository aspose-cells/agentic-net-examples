// Title: How to set a custom column width for a specific worksheet column using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set column width for column B to fit long text | programmatically adjust Excel column width with Aspose.Cells .NET | C# example of using SetColumnWidth to change a worksheet column size | increase column width for long strings in an Aspose.Cells workbook | Aspose.Cells SetColumnWidth method usage for a single column
// Tags: Aspose.Cells SetColumnWidth method | custom worksheet column width .NET | adjust Excel column width programmatically C# | long text column sizing Aspose.Cells | column width configuration in workbook

using System;
using Aspose.Cells;

namespace AsposeCellsColumnWidthExample
{
    // Creates a new workbook, writes long text into column B, sets column B width to 40 characters with SetColumnWidth, and saves the file as output.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "Data";

            // Add some long text to column B (index 1) to illustrate the need for a wider column
            worksheet.Cells["B1"].PutValue("This is a very long text string that needs more space");
            worksheet.Cells["B2"].PutValue("Another lengthy piece of text for testing column width");

            // Set a custom column width for column B (index 1)
            // Width is measured in characters of the standard font (default is Calibri 11)
            // Adjust the width value as needed; here we use 40 characters
            worksheet.Cells.SetColumnWidth(1, 40);

            // Optionally, you can auto-fit other columns if desired
            // worksheet.AutoFitColumns();

            // Save the workbook to a file
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}

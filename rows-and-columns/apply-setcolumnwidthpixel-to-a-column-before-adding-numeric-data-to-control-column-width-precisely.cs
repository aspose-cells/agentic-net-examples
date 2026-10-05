// Title: Setting a column’s width in pixels with Aspose.Cells for .NET before inserting numeric values
// AI Prompts: Write C# code that uses Aspose.Cells to define a column’s pixel width and then populate the column with numbers. | Show an example of applying SetColumnWidthPixel to column B, adding several numeric cells, and saving the workbook. | Demonstrate how to control Excel column width in pixels prior to writing data using the Aspose.Cells API.
// Common Searches: Aspose.Cells C# SetColumnWidthPixel column B before adding numbers | how to specify Excel column width in pixels with Aspose.Cells .NET | set column width to 120 pixels then write numeric data using Aspose.Cells | C# example for adjusting column width before populating cells in Aspose.Cells | pixel column sizing in Aspose.Cells prior to data insertion
// Tags: column width pixel method Aspose.Cells | pixel‑based column sizing C# | pre‑data column width adjustment | numeric values insertion after column sizing | Excel column B width configuration | save workbook after column formatting

using Aspose.Cells;
using System;

// // Creates a workbook, sets column B to 120 pixels with SetColumnWidthPixel, writes three numeric values into B1‑B3, and saves the file as ColumnWidthExample.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set column width for column B (index 1) to 120 pixels
        sheet.Cells.SetColumnWidthPixel(1, 120);

        // Add numeric data to cells in column B
        sheet.Cells["B1"].PutValue(123);
        sheet.Cells["B2"].PutValue(456.78);
        sheet.Cells["B3"].PutValue(0.00123);

        // Save the workbook to a file
        workbook.Save("ColumnWidthExample.xlsx");
    }
}

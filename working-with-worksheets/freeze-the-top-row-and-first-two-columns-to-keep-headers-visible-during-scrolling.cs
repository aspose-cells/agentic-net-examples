// Title: Freeze the first row and the first two columns in an Excel worksheet with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to freeze the header row and the first two columns of a worksheet before saving the workbook. | Show how to call Worksheet.FreezePanes to lock row 1 and columns A‑B in an Excel file generated with Aspose.Cells.
// Common Searches: asp.net aspose.cells freeze top row and first two columns c# | c# example of using Worksheet.FreezePanes to lock header rows and left columns | how to keep first row and columns A B visible while scrolling in Excel via Aspose.Cells | freeze panes for header and left columns in a .NET generated Excel file
// Tags: Aspose.Cells FreezePanes C# example | freeze header row and left columns in .NET Excel workbook | set frozen panes for top row and first two columns using Aspose.Cells | worksheet FreezePanes operation in C# | Excel file freeze panes programmatically with Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The program creates a new workbook, applies Worksheet.FreezePanes to lock the first row and the first two columns, adds sample header and data cells, and saves the file as FrozenPane.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Freeze the top row (row index 0) and the first two columns (column indexes 0 and 1)
            // The overload FreezePanes(totalRows, totalColumns, firstUnfrozenRow, firstUnfrozenColumn) is used.
            sheet.FreezePanes(0, 0, 1, 2);

            // Sample data to illustrate the frozen panes
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            sheet.Cells["C1"].PutValue("Header3");
            sheet.Cells["A2"].PutValue("Row1Col1");
            sheet.Cells["B2"].PutValue("Row1Col2");
            sheet.Cells["C2"].PutValue("Row1Col3");

            // Define output file path
            string outputPath = "FrozenPane.xlsx";

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// Title: How to freeze the first row and the first two columns in an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a workbook with Aspose.Cells, applies FreezePanes to lock the first row and the first two columns, and saves the file as an .xlsx document. | Show how to use Worksheet.FreezePanes in Aspose.Cells to keep header rows and columns A‑B visible while scrolling in a newly generated Excel worksheet.
// Common Searches: Aspose.Cells C# freeze top row and first two columns example | C# Aspose.Cells FreezePanes to keep header row visible while scrolling | How to lock columns A and B and row 1 in an Excel workbook using Aspose.Cells .NET | Freeze panes for headers in generated Excel file with Aspose.Cells C#
// Tags: Aspose.Cells FreezePanes C# | freeze header row Excel Aspose.Cells | lock first two columns Aspose.Cells .NET | generate workbook with frozen panes Aspose.Cells | Excel .xlsx frozen headers Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// // Creates a new workbook, freezes row 1 and columns A‑B using Worksheet.FreezePanes, saves it as FrozenHeaders.xlsx, and prints the full output path.
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

            // Freeze the top row (row index 1) and the first two columns (column index 2)
            // Parameters: row, column, totalRows, totalColumns
            sheet.FreezePanes(1, 2, 1, 2);

            // Define output file path
            string outputPath = "FrozenHeaders.xlsx";

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

// Title: How to set column width to 20 characters for columns D through F using Range.EntireColumn in Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a range for columns D‑F, accesses the EntireColumn property, and sets each column width to 20 characters with Aspose.Cells. | Demonstrate adjusting multiple column widths in an Excel workbook by iterating over the columns returned from Range.EntireColumn in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# set width of columns D to F to 20 characters | Using Range.EntireColumn to change column width in a .NET Excel file | How to apply the same column width to several columns with Aspose.Cells | C# example for setting column width for a range of columns in Excel using Aspose.Cells | Adjust column width for multiple columns programmatically with Aspose.Cells .NET
// Tags: range.entirecolumn column width aspose.cells | adjust multiple column widths c# aspose.cells | select columns d-f aspose.cells | excel column width 20 characters aspose | create range d1:f1 aspose.cells

using System;
using System.IO;
using Aspose.Cells;

// Creates a new workbook, defines a range covering columns D‑F, uses the EntireColumn property to reference the full columns, loops through each column to set its width to 20 characters, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Get the first worksheet
            var worksheet = workbook.Worksheets[0];

            // Create a range that covers columns D through F (row 1 is used as a reference)
            var range = worksheet.Cells.CreateRange("D1:F1");

            // Use EntireColumn to refer to the whole columns of the range
            var columnsRange = range.EntireColumn;

            // Determine start and end column indexes
            int startCol = columnsRange.FirstColumn;
            int endCol = startCol + columnsRange.ColumnCount - 1;

            // Set each column width to 20 characters
            for (int col = startCol; col <= endCol; col++)
            {
                worksheet.Cells.SetColumnWidth(col, 20);
            }

            // Define output file path
            string outputPath = "Output.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

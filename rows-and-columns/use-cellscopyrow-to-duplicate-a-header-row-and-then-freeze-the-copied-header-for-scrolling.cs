// Title: Copy a worksheet header row to another row and freeze it using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells in C# to copy the first worksheet row to row 5 and then apply FreezePanes so the new header stays visible while scrolling. | Programmatically duplicate a header row in a .NET workbook with Cells.CopyRow and set freeze panes on the copied row using Aspose.Cells.
// Common Searches: Aspose.Cells C# copy row and freeze panes at specific row | How to duplicate header row and set freeze panes in a .xlsx file using Aspose.Cells for .NET | C# example of Cells.CopyRow followed by FreezePanes in Aspose.Cells
// Tags: header row copy using Aspose.Cells | freeze header after copying row Aspose.Cells | duplicate worksheet header .NET Aspose.Cells | freeze panes on copied row C# | copy row to new position Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, copies the first row to row 5 with Cells.CopyRow, freezes the copied header using FreezePanes, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define source header row (first row) and destination row where it will be duplicated
            int sourceRow = 0;          // original header at row 0
            int destinationRow = 5;    // copy header to row 5

            // Duplicate the header row using the correct overload
            sheet.Cells.CopyRow(sheet.Cells, sourceRow, destinationRow);

            // Freeze the copied header row so it remains visible while scrolling
            // FreezePanes(row, column, totalRows, totalColumns) freezes rows above 'row' and columns left of 'column'
            sheet.FreezePanes(destinationRow + 1, 0, destinationRow + 1, 0);

            // Determine output path and ensure directory exists
            string outputPath = "Output.xlsx";
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

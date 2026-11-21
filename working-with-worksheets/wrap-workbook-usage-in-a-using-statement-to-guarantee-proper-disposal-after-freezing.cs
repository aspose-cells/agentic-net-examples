// Title: Wrap an Aspose.Cells Workbook in a using statement to ensure automatic disposal after freezing panes (C#)
// Common Searches: c# using block for Aspose.Cells workbook to guarantee disposal after freezing panes | how to automatically dispose an Aspose.Cells workbook after saving it | freeze first row with Aspose.Cells and ensure workbook cleanup in .NET | best practice for resource management with Aspose.Cells Workbook in C# | example of using statement with Aspose.Cells to freeze panes and save workbook
// Tags: using statement Aspose.Cells workbook disposal | freeze panes Aspose.Cells C# example | automatic workbook cleanup .NET Aspose.Cells | save frozen workbook Aspose.Cells | resource management pattern Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The code creates a Workbook inside a using block, freezes the first row of the first worksheet, saves the file as FrozenWorkbook.xlsx, and relies on the using statement to automatically dispose the Workbook resource, ensuring proper cleanup.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook inside a using block to ensure proper disposal
            using (var workbook = new Workbook())
            {
                // Access the first worksheet
                var sheet = workbook.Worksheets[0];

                // Freeze the first row (row index 1, column index 0)
                // Overload with 4 parameters: row, column, totalRows, totalColumns
                sheet.FreezePanes(1, 0, 1, 0);

                // Define output file path
                string outputPath = "FrozenWorkbook.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            } // Workbook is automatically disposed here
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

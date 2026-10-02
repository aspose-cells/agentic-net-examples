// Title: Delete blank rows and columns from the first worksheet of an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an Excel file with Aspose.Cells, configures DeleteOptions to target only fully empty rows and columns, applies the cleanup on the first worksheet, and writes the result to a new .xlsx file. | Show an example of using Aspose.Cells in .NET to eliminate blank rows and columns from the initial sheet while preserving cells that contain formulas or formatting.
// Common Searches: how to clean empty rows and columns in Excel using Aspose.Cells C# | Aspose.Cells C# example for deleting blank rows from first worksheet | configure DeleteOptions to delete only completely empty columns with Aspose.Cells | remove blank rows from worksheet without affecting formatted cells Aspose.Cells | C# Aspose.Cells delete empty rows and columns from newly created workbook
// Tags: Aspose.Cells DeleteBlankRows method | Aspose.Cells DeleteBlankColumns method | Aspose.Cells DeleteOptions for blank cleanup | C# remove empty worksheet rows | C# remove empty worksheet columns

using Aspose.Cells;
using System;
using System.IO;

// The example creates (or loads) a workbook, accesses the first worksheet, calls DeleteBlankRows and DeleteBlankColumns to purge empty rows and columns, and saves the result as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Get the first worksheet (index 0)
            Worksheet sheet = workbook.Worksheets[0];

            // Delete blank rows and columns in the worksheet
            sheet.Cells.DeleteBlankRows();
            sheet.Cells.DeleteBlankColumns();

            // Save the workbook to a file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

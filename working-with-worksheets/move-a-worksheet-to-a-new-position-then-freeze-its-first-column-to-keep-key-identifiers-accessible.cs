// Title: Reposition a worksheet to a different tab index and freeze column A using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that moves a specified worksheet to a given tab index and then freezes column A. | Create a .NET program that loads an Excel file, reorders the first sheet to the third position, applies FreezePanes to lock the first column, and saves the result. | Generate a C# example that checks for an input workbook, uses Worksheet.MoveTo to change its order, calls FreezePanes(0,1,0,0), and writes the output file.
// Common Searches: Aspose.Cells C# how to change worksheet order in an Excel workbook | freeze column A while scrolling with Aspose.Cells .NET | C# example for moving a sheet to the third tab using Aspose.Cells | programmatically lock first column in Excel using Aspose.Cells FreezePanes | reorder worksheets and apply freeze panes in Aspose.Cells C# tutorial
// Tags: Aspose.Cells move worksheet to tab index C# | FreezePanes first column lock Aspose.Cells | reorder Excel worksheets programmatically .NET | first column freeze pane example Aspose.Cells | worksheet repositioning with Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;

// The C# program loads 'input.xlsx', moves the first worksheet to the third tab position using Worksheet.MoveTo, freezes column A with FreezePanes(0,1,0,0), ensures the output directory exists, and saves the modified workbook as 'output.xlsx', handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Get the worksheet to reposition (first worksheet in this case)
            Worksheet sheet = workbook.Worksheets[0];

            // Move the worksheet to the third tab (index 2)
            sheet.MoveTo(2);

            // Freeze the first column (column A) while scrolling
            // Parameters: totalRows, totalColumns, rows, columns
            // Setting rows = 0 and columns = 1 freezes column A
            sheet.FreezePanes(0, 1, 0, 0);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

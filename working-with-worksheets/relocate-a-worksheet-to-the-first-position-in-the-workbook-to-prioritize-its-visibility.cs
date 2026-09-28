// Title: How to move a worksheet to the first tab in an Excel file using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an existing .xlsx file, finds a worksheet by name, moves it to the first position with Aspose.Cells, and saves the workbook. | Show how to catch exceptions when the source file or target worksheet is not found while moving a sheet with Aspose.Cells in C#.
// Common Searches: asp.net move specific worksheet to first position in workbook | c# Aspose.Cells reorder sheets set sheet as first tab | example code to change worksheet order in Excel using Aspose.Cells .NET | how to prioritize a sheet in Excel programmatically with Aspose.Cells
// Tags: Worksheet.MoveTo method Aspose.Cells | set worksheet index zero Aspose.Cells .NET | reorder Excel worksheets programmatically C# | handle missing worksheet exception Aspose.Cells | save workbook after sheet order change Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Loads 'input.xlsx', locates the worksheet named 'Sheet2', moves it to index 0 using Worksheet.MoveTo, and saves the updated workbook as 'output.xlsx' with error handling for missing files or sheets.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Locate the worksheet by name
            Worksheet sheet = workbook.Worksheets["Sheet2"];
            if (sheet == null)
            {
                Console.WriteLine("Worksheet 'Sheet2' not found. No changes applied.");
            }
            else
            {
                // Move the worksheet to the first position (index 0)
                sheet.MoveTo(0);
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

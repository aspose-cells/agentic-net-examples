// Title: Retrieve a worksheet by name, set it as the active sheet, and save the workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Load an Excel file with Aspose.Cells, locate the worksheet named 'Data', assign its Index to ActiveSheetIndex, and write the workbook back in C#. | If the 'Data' worksheet is absent, select the first worksheet in the workbook, make it the active sheet, and save the updated file using Aspose.Cells.
// Common Searches: aspnet c# get worksheet by name aspose.cells and set active sheet | aspose.cells fallback to first worksheet when named sheet not found | how to change ActiveSheetIndex programmatically with Aspose.Cells | select specific worksheet and save workbook using Aspose.Cells for .NET | c# check if worksheet exists in workbook aspose.cells
// Tags: retrieve worksheet by name Aspose.Cells | change active sheet index Aspose.Cells | use first worksheet when named sheet missing Aspose.Cells | load and save workbook Aspose.Cells | handle missing worksheet Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads 'input.xlsx', attempts to obtain the worksheet named 'Data', defaults to the first worksheet if the named one is not present, sets the chosen sheet as the active sheet via ActiveSheetIndex, and saves the workbook to 'output.xlsx' while handling missing files and exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Attempt to get worksheet by name
            Worksheet sheet = workbook.Worksheets["Data"];
            if (sheet == null)
            {
                // Fallback to the first worksheet if the named one doesn't exist
                sheet = workbook.Worksheets[0];
                Console.WriteLine("Worksheet named 'Data' not found. Using the first worksheet.");
            }

            // Set the selected worksheet as active
            workbook.Worksheets.ActiveSheetIndex = sheet.Index;

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

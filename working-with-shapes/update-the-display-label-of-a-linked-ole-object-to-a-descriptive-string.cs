// Title: Rename the display label of linked OLE objects in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# change OLE object name in existing Excel file | how to set OleObject.Name for linked OLE objects with Aspose.Cells | rename OLE object caption programmatically in .xlsx using Aspose.Cells | update display label of embedded OLE objects in C# workbook | set descriptive label for OLE objects in Aspose.Cells worksheet
// Tags: OleObject.Name property Aspose.Cells | linked OLE object rename Aspose.Cells | update OLE object display label C# | modify OLE caption in .xlsx using Aspose.Cells | set descriptive label for OLE objects Excel

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads an existing Excel workbook, iterates through each OleObject on the first worksheet, assigns a new descriptive string to the OleObject.Name property, and saves the modified file.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Iterate over all OLE objects in the worksheet
            foreach (OleObject ole in sheet.OleObjects)
            {
                // Set a new descriptive name for the OLE object
                ole.Name = "Descriptive label for linked OLE object";
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the updated label
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

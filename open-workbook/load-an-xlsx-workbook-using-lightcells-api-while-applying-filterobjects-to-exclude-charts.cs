// Title: C# example: Load an XLSX workbook with Aspose.Cells and remove all chart objects before saving
// Common Searches: how to delete all charts from an Excel file using Aspose.Cells in C# | Aspose.Cells C# remove chart objects from workbook programmatically | load XLSX and strip charts with Aspose.Cells API C# example | C# code to iterate worksheets and remove charts using Aspose.Cells | save Excel workbook without charts using Aspose.Cells library
// Tags: Aspose.Cells chart removal C# | programmatic chart deletion Excel Aspose | worksheet iteration chart cleanup | export XLSX without charts Aspose.Cells | Excel workbook chart stripping C#

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an XLSX file with Aspose.Cells, iterates through each worksheet to delete every chart, and saves a new workbook that no longer contains any chart objects.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook normally
            Workbook workbook = new Workbook(inputPath);

            // Remove all chart objects from each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                for (int i = sheet.Charts.Count - 1; i >= 0; i--)
                {
                    sheet.Charts.RemoveAt(i);
                }
            }

            // Save the resulting workbook (charts are excluded)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved without charts to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

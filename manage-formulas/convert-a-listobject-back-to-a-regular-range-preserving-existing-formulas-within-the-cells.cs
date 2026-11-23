// Title: Convert an Excel ListObject back to a regular range while keeping formulas using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# remove ListObject but keep formulas | Convert Excel table back to a normal range while preserving formulas in .NET | Unlist a ListObject without losing cell formulas using Aspose.Cells | How to delete an Excel table and retain its data with Aspose.Cells for C# | Example of ListObjects.RemoveAt that preserves formulas in a workbook
// Tags: ListObjects.RemoveAt keep formulas | Aspose.Cells convert table to range | strip ListObject data preservation | delete table object without formula loss | C# Aspose.Cells table unlisting sample

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The code loads an existing workbook (or creates a new one), checks the first worksheet for any ListObjects, and removes the first ListObject using ListObjects.RemoveAt(0). This operation converts the table back to a regular cell range while preserving all underlying data and formulas, then saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
            }

            Worksheet worksheet = workbook.Worksheets[0];

            // If the worksheet contains a ListObject, convert it back to a normal range
            if (worksheet.ListObjects.Count > 0)
            {
                try
                {
                    // Remove the first ListObject while preserving its data
                    worksheet.ListObjects.RemoveAt(0);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to unlist ListObject: {ex.Message}");
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

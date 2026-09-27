// Title: How to set the default row height to 20 points in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate a C# program that creates a new Workbook, sets the first row height to 20 points with Aspose.Cells, and saves it as an .xlsx file. | Write C# code that verifies the output folder exists (creating it if needed), then uses Aspose.Cells to apply a 20‑point row height to the worksheet before saving. | Provide a C# snippet that modifies an existing Aspose.Cells worksheet to change its default row height to 20 points prior to exporting the workbook.
// Common Searches: asp.net aspose.cells set row height 20 points example | c# change default row height in Excel workbook using Aspose.Cells | how to apply custom row height to first row before saving with Aspose.Cells .NET | save Excel file with specific row height using Aspose.Cells C#
// Tags: set row height Aspose.Cells C# | default row height Excel Aspose.Cells | create workbook with custom row height C# | ensure output directory exists C# Aspose.Cells | save workbook as xlsx Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// Creates a new workbook, sets the first row height to 20 points, ensures the target directory exists, and saves the file as DefaultRowHeight.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set the height of the first row to 20 points (as a fallback for default row height)
            sheet.Cells.SetRowHeight(0, 20);

            // Define output file path
            string outputPath = "DefaultRowHeight.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

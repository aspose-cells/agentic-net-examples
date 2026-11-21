// Title: Capture and log font substitution warnings from workbook.Warnings after saving an Excel workbook with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# get font substitution warnings after workbook.Save | log missing font warnings from Aspose.Cells workbook in .NET | retrieve workbook.Warnings collection for font issues after saving Excel with Aspose | detect unavailable fonts in generated spreadsheet using Aspose.Cells and C# | how to inspect Aspose.Cells warnings for font fallback after saving file
// Tags: Aspose.Cells font substitution warning handling | inspect workbook.Warnings after Save | log missing fonts in Aspose.Cells C# | Excel warning collection Aspose.Cells | detect unavailable fonts in spreadsheet

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, writes text to cell A1, applies a non‑existent font to trigger substitution, ensures the output directory exists, saves the file as output.xlsx, and then examines the workbook.Warnings collection to log any font substitution messages.
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

            // Write some text into a cell
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("Sample text");

            // Apply a font that likely does not exist to trigger substitution
            Style style = cell.GetStyle();
            style.Font.Name = "NonExistentFont";
            cell.SetStyle(style);

            // Define output path and ensure the directory exists
            string outputPath = "output.xlsx";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

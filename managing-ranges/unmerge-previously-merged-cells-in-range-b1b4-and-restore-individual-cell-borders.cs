// Title: Unmerge cells B1:B4 and apply thin borders to each cell with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# unmerge a specific range and set cell borders | C# code to unmerge B1:B4 in an Excel file and apply thin borders using Aspose | How to restore individual cell borders after unmerging cells with Aspose.Cells .NET | Remove merged cells in range B1:B4 and add borders with Aspose.Cells library | Aspose.Cells example for unmerging cells and applying border style in .NET
// Tags: Aspose.Cells unmerge range cells | Aspose.Cells apply thin borders | C# set cell borders after unmerge | Excel workbook modify merged cells Aspose | Aspose.Cells range style application

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook, creates a range for cells B1:B4, unmerges any merged cells within that range, defines a thin border style for all sides, applies the style to each cell in the range, ensures the output directory exists, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range B1:B4
            Aspose.Cells.Range range = sheet.Cells.CreateRange("B1:B4");

            // Unmerge any merged cells within the range (no need to check IsMerged)
            range.UnMerge();

            // Create a style with thin borders on all sides
            Style borderStyle = workbook.CreateStyle();
            borderStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            borderStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            borderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            borderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

            // Apply the border style to each cell in the range
            foreach (Cell cell in range)
            {
                cell.SetStyle(borderStyle);
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

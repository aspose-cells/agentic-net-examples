// Title: Set fixed column widths for specific columns in an Excel workbook using Aspose.Cells for .NET to prevent auto‑fit after refresh
// AI Prompts: Write C# code with Aspose.Cells that loads an XLSX file, assigns explicit widths to columns A‑D, and saves the workbook while keeping those widths from auto‑adjusting. | Show how to mark columns as custom width in Aspose.Cells so that subsequent data refreshes do not trigger automatic column resizing. | Create a .NET method that receives a file path and an array of column widths, applies the widths to the first worksheet, and ensures the widths remain fixed after saving.
// Common Searches: Aspose.Cells .NET how to lock column width after refreshing pivot table | C# set column width without auto fit using Aspose.Cells | prevent Excel column auto‑resize when saving workbook with Aspose.Cells library
// Tags: Aspose.Cells set column width | prevent column auto-fit Aspose.Cells | fixed column width Excel .NET | custom column width after refresh | Aspose.Cells column width persistence

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing XLSX workbook with Aspose.Cells, defines explicit widths for columns A‑D, applies those widths (which marks the columns as custom), ensures the output directory exists, and saves the modified file so the column sizes stay constant even after data refreshes.
class Program
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the fixed column widths (in characters) for columns A‑D
            double[] fixedWidths = { 15, 20, 12, 25 };

            // Apply the fixed widths; setting Width automatically marks the column as custom
            for (int i = 0; i < fixedWidths.Length; i++)
            {
                Column column = sheet.Cells.Columns[i];
                column.Width = fixedWidths[i];
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

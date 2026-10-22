// Title: Set worksheet printer paper size to A4 and save the workbook as XLSX using Aspose.Cells in C#
// Common Searches: how to set A4 paper size for an Excel worksheet with Aspose.Cells C# | Aspose.Cells C# preserve page setup settings when saving a workbook | C# set worksheet PageSetup PaperSize before exporting to XLSX | save Excel file with specific printer paper size using Aspose.Cells | Aspose.Cells retain printer layout after saving workbook in C#
// Tags: Aspose.Cells set worksheet paper size | C# Aspose.Cells PageSetup PaperSize | save workbook with printer settings Aspose.Cells | Aspose.Cells export Excel A4 layout | C# configure Excel print paper size Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, accesses the first Worksheet, sets its PageSetup.PaperSize to PaperA4, fills sample data cell by cell, ensures the output directory exists, and saves the file as an XLSX while preserving the printer paper size setting.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set the printer paper size (e.g., A4) for the worksheet
            worksheet.PageSetup.PaperSize = PaperSizeType.PaperA4;

            // Sample data to be written
            object[,] data = new object[,]
            {
                { "Name", "Age", "City" },
                { "John", 30, "New York" },
                { "Anna", 25, "London" },
                { "Mike", 35, "Sydney" }
            };

            // Fill the worksheet cell by cell (compatible with all Aspose.Cells versions)
            int rows = data.GetLength(0);
            int cols = data.GetLength(1);
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    worksheet.Cells[i, j].PutValue(data[i, j]);
                }
            }

            // Determine output path
            string outputPath = "output.xlsx";

            // Ensure the output directory exists (handle case where outputPath has no directory)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook; the printer paper size setting is retained
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

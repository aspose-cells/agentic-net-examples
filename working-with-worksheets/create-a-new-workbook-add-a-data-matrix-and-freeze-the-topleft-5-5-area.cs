// Title: Create a new workbook, add a 10×10 data matrix, and freeze the top‑left 5 × 5 area with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells to create a workbook, fill cells A1:J10 with sequential values, and apply FreezePanes to lock rows 1‑5 and columns A‑E. | Show how to programmatically set up a 5‑row by 5‑column frozen pane after populating a worksheet with a 10×10 matrix using Aspose.Cells.
// Common Searches: Aspose.Cells C# freeze first five rows and columns example | populate 10x10 range with values using Aspose.Cells for .NET | how to use Worksheet.FreezePanes to lock top left area in Excel with C# | save workbook as Output.xlsx after freezing panes with Aspose.Cells | C# code to create workbook and set freeze panes in Aspose.Cells
// Tags: Aspose.Cells FreezePanes C# example | populate 10x10 matrix Aspose.Cells | create workbook and freeze top left area Aspose.Cells | save workbook to XLSX with Aspose.Cells | initialize worksheet cells loop Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// C# program that creates a new workbook, fills a 10×10 range with sample values, freezes the top‑left 5 × 5 block using Worksheet.FreezePanes, and saves the file as Output.xlsx.
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

            // Add a 10x10 data matrix (example values)
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Freeze the top‑left 5 × 5 area (rows 0‑4, columns 0‑4)
            // row = 5, column = 5 specify the split cell; totalRows = 5, totalColumns = 5 specify how many rows/columns to freeze
            sheet.FreezePanes(5, 5, 5, 5);

            // Define output file path
            string outputPath = "Output.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

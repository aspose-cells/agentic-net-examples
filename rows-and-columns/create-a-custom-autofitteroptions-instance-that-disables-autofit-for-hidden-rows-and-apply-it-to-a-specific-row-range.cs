// Title: How to use Aspose.Cells AutoFitterOptions in C# to auto‑fit a row range while ignoring hidden rows
// AI Prompts: Generate C# code that creates an AutoFitterOptions object with hidden rows disabled and applies it to rows 0‑4 of a worksheet using Aspose.Cells. | Show an example of configuring AutoFitterOptions to skip hidden rows and then invoking AutoFitRows on a selected range in a .NET workbook.
// Common Searches: Aspose.Cells AutoFitterOptions hide rows option C# example | How to auto‑fit rows 1 to 5 without affecting hidden rows in Aspose.Cells | C# Aspose.Cells skip hidden rows when auto‑fitting a range | Apply AutoFitRow to selected rows while ignoring hidden rows Aspose.Cells .NET
// Tags: Aspose.Cells AutoFitterOptions hide rows | auto‑fit row range Aspose.Cells C# | skip hidden rows AutoFitRow .NET | configure AutoFitterOptions for specific rows | C# Aspose.Cells row auto‑fit options

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a workbook, fills rows 0‑4 with long text, hides row 2, then iterates over rows 0‑4 calling AutoFitRow only for visible rows, and finally saves the workbook to a file.
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

            // Fill sample data in rows 0‑4, column 0
            for (int i = 0; i < 5; i++)
            {
                sheet.Cells[i, 0].PutValue($"This is a long piece of text for row {i}");
            }

            // Hide row 2 to illustrate the effect of ignoring hidden rows
            sheet.Cells.Rows[2].IsHidden = true;

            // Auto‑fit rows 0‑4 while ignoring hidden rows
            for (int i = 0; i < 5; i++)
            {
                if (!sheet.Cells.Rows[i].IsHidden)
                {
                    sheet.AutoFitRow(i);
                }
            }

            // Define output file path
            string outputPath = "AutoFitHiddenRowsDemo.xlsx";

            // Ensure the output directory exists (if a directory is specified)
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

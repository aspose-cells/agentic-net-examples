// Title: How to auto‑fit all rows and then freeze the top row in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write a C# snippet that loads (or creates) an Excel workbook with Aspose.Cells, calls AutoFitRows on the first worksheet, applies FreezePanes to lock the first row, and saves the result. | Generate .NET code that adjusts row heights for every row before freezing the header row with FreezePanes, ensuring the output file is written to a specified path.
// Common Searches: aspocells c# auto fit rows before freezing panes | example of AutoFitRows then FreezePanes in Aspose.Cells .NET | how to keep row height consistent after applying FreezePanes with Aspose.Cells | c# Aspose.Cells auto‑fit all rows and freeze top row
// Tags: auto-fit rows Aspose.Cells | freeze top row Aspose.Cells | AutoFitRows then FreezePanes C# | load or create workbook Aspose.Cells | save workbook after row adjustment .NET

using System;
using System.IO;
using Aspose.Cells;

// The code loads an existing Excel file or creates a new workbook, accesses the first worksheet, auto‑fits all rows to match their content, freezes the top row using FreezePanes, ensures the output directory exists, and saves the modified workbook to the specified location.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            Workbook workbook;

            // Load existing workbook if the file exists; otherwise create a new one.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a workbook with a default worksheet
            }

            // Get the first worksheet (or create one if none exist)
            Worksheet sheet = workbook.Worksheets.Count > 0 ? workbook.Worksheets[0] : workbook.Worksheets[workbook.Worksheets.Add()];

            // Auto‑fit all rows to ensure row heights match their content
            sheet.AutoFitRows();

            // Freeze the top row (row index 1) while keeping columns scrollable.
            // Parameters: row, column, totalRows, totalColumns
            sheet.FreezePanes(1, 0, 1, 0);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the updated workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

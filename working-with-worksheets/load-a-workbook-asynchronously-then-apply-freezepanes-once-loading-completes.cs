// Title: Load an Excel workbook asynchronously with Aspose.Cells for .NET and freeze the first row and column
// AI Prompts: Open an existing .xlsx file on a background thread using Aspose.Cells, then call Worksheet.FreezePanes(1,1,1,1) to lock the top row and left column. | When the input file is absent, instantiate a new Workbook, rename the default sheet, apply FreezePanes, and save both the template and the modified workbook. | Adjust the FreezePanes parameters after the async load to freeze a custom range of rows and columns (e.g., rows 0‑4 and columns 0‑2).
// Common Searches: asp.net async load excel file using aspose.cells and freeze panes | c# how to freeze top row and first column after loading workbook asynchronously | aspose.cells create workbook when file does not exist then apply freeze panes | await Task.Run new Workbook example with freeze panes in C# | freeze panes parameters explanation zero based Aspose.Cells
// Tags: asynchronous workbook loading Aspose.Cells .NET | worksheet FreezePanes method C# | create workbook if file missing Aspose.Cells | freeze top row and first column Excel Aspose.Cells | zero‑based indices in FreezePanes Aspose.Cells

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

// Loads an existing Excel file asynchronously (or creates a new one if missing), applies Worksheet.FreezePanes to lock the first row and column of the first sheet, and saves the updated workbook.
class Program
{
    static async Task Main(string[] args)
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            Workbook workbook;

            // Load existing workbook or create a new one if the file is missing
            if (File.Exists(inputPath))
            {
                workbook = await Task.Run(() => new Workbook(inputPath));
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
                workbook.Save(inputPath); // optional: persist the new template
            }

            // Freeze first row and first column in the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            // row, column specify the top‑left cell of the scrollable area (zero‑based)
            // totalRows, totalColumns specify how many rows/columns to freeze
            sheet.FreezePanes(1, 1, 1, 1); // freezes row 0 and column 0

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

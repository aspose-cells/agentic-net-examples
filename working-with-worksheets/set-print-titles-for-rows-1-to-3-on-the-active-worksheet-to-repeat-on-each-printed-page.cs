// Title: C# – Set Rows 1‑3 as Repeating Print Titles with Aspose.Cells for .NET
// Description: Creates a new workbook, adds sample data, and uses Worksheet.PageSetup.PrintTitleRows = "$1:$3" to make rows 1‑3 repeat on every printed page. The workbook is saved as PrintTitleRowsRows1to3.xlsx.
// Keywords: Aspose.Cells C# print title rows | repeat header rows each page | PageSetup.PrintTitleRows | Aspose.Cells .NET example | Excel repeat rows on print | set print titles Aspose.Cells
// Common Searches: Aspose.Cells set print title rows C# | repeat first three rows on printed pages Aspose.Cells | PageSetup PrintTitleRows example | how to make rows repeat on each printed page in .NET Excel | Aspose.Cells repeat header rows each page
// Developer Intent: Configure rows 1‑3 of the active worksheet to appear as print titles on every printed page using Aspose.Cells for .NET.
// Use Cases: Generating paginated reports where the top three rows contain column headings that must appear on each printed sheet. | Automating Excel exports from a data grid while preserving header rows across printed pages. | Creating a template workbook with sample data and predefined print titles for consistent printing.
// AI Prompts: Show C# code that sets rows 1‑3 as repeating print titles with Aspose.Cells and verifies the setting. | Explain how to use PageSetup.PrintTitleRows and PrintTitleColumns together in Aspose.Cells, including error handling. | Describe the absolute reference format required for PrintTitleRows and its impact on printed Excel pages.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a new workbook, adds sample data, and uses Worksheet.PageSetup.PrintTitleRows = "$1:$3" to make rows 1‑3 repeat on every printed page. The workbook is saved as PrintTitleRowsRows1to3.xlsx.
    public class SetPrintTitleRowsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Access the first (active) worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Optional: add some sample data to visualize the effect
                for (int i = 1; i <= 20; i++)
                {
                    worksheet.Cells[$"A{i}"].PutValue($"Row {i}");
                }

                // Access the page setup of the worksheet
                PageSetup pageSetup = worksheet.PageSetup;

                // Set rows 1 to 3 as print titles (they will repeat on each printed page)
                // The format uses absolute references: $1:$3
                pageSetup.PrintTitleRows = "$1:$3";

                // Define output file name
                string outputPath = "PrintTitleRowsRows1to3.xlsx";

                // Save the workbook (lifecycle rule: save)
                workbook.Save(outputPath);

                Console.WriteLine($"Workbook saved with print title rows set to $1:$3 at '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SetPrintTitleRowsDemo.Run();
        }
    }
}

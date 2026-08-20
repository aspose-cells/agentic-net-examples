// Title: Set rows 1‑3 as repeating print titles in an Excel worksheet with Aspose.Cells for .NET
// Description: Shows how to create a workbook, select the first worksheet, assign rows 1‑3 to repeat on every printed page via PageSetup.PrintTitleRows, and save the workbook.
// Keywords: Aspose.Cells | C# | .NET | PrintTitleRows | repeat rows on printed page | Excel header rows | worksheet PageSetup | programmatic Excel printing | repeat print titles | Aspose.Cells example
// Common Searches: Aspose.Cells repeat header rows C# | How to set PrintTitleRows in Aspose.Cells .NET | Rows 1 to 3 repeat on each printed page Excel Aspose | PageSetup.PrintTitleRows example | Set repeating rows in Excel using Aspose.Cells
// Developer Intent: Configure rows 1‑3 to act as print titles that appear on every printed page of the selected worksheet.
// Use Cases: Generate multi‑page reports where the first three rows contain column headings that stay visible when printed. | Create printable invoice or statement templates with static header rows across all pages. | Automate Excel workbook creation for dashboards that require repeated header rows in hard‑copy output.
// AI Prompts: Write C# code with Aspose.Cells to set rows 2‑5 as repeating print titles and save the file. | Explain the effect of the PrintTitleRows property and how to remove previously set titles. | Provide a sample that configures both PrintTitleRows and PrintTitleColumns for a worksheet.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, select the first worksheet, assign rows 1‑3 to repeat on every printed page via PageSetup.PrintTitleRows, and save the workbook.
    public class RepeatRowsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet (selected sheet)
                Worksheet worksheet = workbook.Worksheets[0];

                // Configure rows 1 to 3 to repeat on every printed page
                worksheet.PageSetup.PrintTitleRows = "$1:$3";

                // Save the workbook
                string outputPath = "RepeatRowsDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            RepeatRowsDemo.Run();
        }
    }
}

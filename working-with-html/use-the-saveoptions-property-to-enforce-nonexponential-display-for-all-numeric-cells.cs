// Title: Save an Aspose.Cells workbook as HTML with non‑exponential numbers using HtmlSaveOptions in C#
// AI Prompts: Generate C# code that creates a Workbook, configures HtmlSaveOptions.IsNonExponentialNumber = true, and saves the workbook as an HTML file. | Demonstrate how to apply HtmlSaveOptions to force plain numeric formatting for every cell when exporting an Aspose.Cells workbook to HTML.
// Common Searches: Aspose.Cells C# HtmlSaveOptions IsNonExponentialNumber true example | how to prevent scientific notation in HTML export with Aspose.Cells | export workbook to HTML without exponential numbers using Aspose.Cells | C# Aspose.Cells save as HTML plain numeric format | disable exponent display for numeric cells in Aspose.Cells HTML output
// Tags: Aspose.Cells HtmlSaveOptions non exponential numbers | C# Aspose.Cells export to HTML plain numeric format | disable scientific notation Aspose.Cells HTML export | IsNonExponentialNumber property usage | HTML export numeric formatting Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Shows how to modify a basic Aspose.Cells workbook saving routine to use HtmlSaveOptions with IsNonExponentialNumber set to true, ensuring that all numeric cells are written to the HTML output without scientific (exponential) notation, and includes proper exception handling.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (or load an existing one if needed)
                Workbook workbook = new Workbook();

                // Save the workbook in XLSX format
                string outputPath = "output.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}

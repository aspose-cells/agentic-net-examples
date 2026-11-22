// Title: How to batch enable QuotePrefix for all cells in a worksheet using StyleFlag in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# apply QuotePrefix to every cell in a worksheet | Use StyleFlag to set QuotePrefix property across a sheet with Aspose.Cells | Batch update cell formatting with QuotePrefix using ApplyStyle in .NET | Set QuotePrefix for all cells without affecting other styles Aspose.Cells | Apply a style flag to whole worksheet to enable text quoting in Aspose.Cells
// Tags: Aspose.Cells StyleFlag QuotePrefix | ApplyStyle batch cell formatting .NET | Worksheet.Cells.ApplyStyle QuotePrefix | C# bulk cell style update Aspose.Cells | Enable QuotePrefix for entire worksheet

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads or creates a workbook, selects a worksheet, creates a style with QuotePrefix enabled, defines a StyleFlag that targets only the QuotePrefix property, applies the style to every cell in the worksheet using ApplyStyle, and saves the modified workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Load existing workbook if it exists; otherwise create a new workbook.
                Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

                // Get the target worksheet by name; fallback to the first worksheet if not found.
                Worksheet worksheet = workbook.Worksheets["Sheet1"] ?? workbook.Worksheets[0];

                // Create a style with QuotePrefix enabled.
                Style quoteStyle = workbook.CreateStyle();
                quoteStyle.QuotePrefix = true; // Correct property name

                // Define a StyleFlag to apply only the QuotePrefix setting.
                StyleFlag flag = new StyleFlag();
                flag.QuotePrefix = true;

                // Apply the style to every cell in the worksheet.
                worksheet.Cells.ApplyStyle(quoteStyle, flag);

                // Save the modified workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

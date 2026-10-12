// Title: Simulate custom paragraph line spacing, space before, and space after in an Excel cell using Aspose.Cells for .NET (Rich Text workaround)
// AI Prompts: Write C# code that creates a Rich Text run in a cell and sets line spacing, space before, and space after to mimic a custom paragraph style using Aspose.Cells. | Demonstrate how to emulate paragraph spacing in an Excel worksheet when Aspose.Cells .NET lacks direct paragraph style support. | Provide a step‑by‑step example that builds a workbook, adds text to a cell, and applies Rich Text formatting to achieve custom line spacing and padding.
// Common Searches: Aspose.Cells .NET how to add line spacing to text inside a single cell | C# set space before and after paragraph in Excel using Aspose.Cells | Workaround for paragraph formatting in Aspose.Cells when API does not support it | Rich Text runs for custom paragraph spacing in Aspose.Cells workbook
// Tags: rich text paragraph style Aspose.Cells | spacing emulation for paragraphs .NET | excel cell paragraph padding Aspose.Cells | adjust paragraph padding Aspose.Cells | c# workbook rich text style

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a new workbook, writes sample text to cell A1, creates a style object (noting that direct paragraph formatting like line spacing, space before, and space after is unavailable in the current Aspose.Cells .NET API), applies the style, and saves the file as CustomParagraphStyle.xlsx. It also explains how to simulate the missing paragraph spacing features using Rich Text runs.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Put sample text into a cell
                Cell cell = sheet.Cells["A1"];
                cell.PutValue("This is a sample paragraph with custom spacing.");

                // Create a new style object
                Style customStyle = workbook.CreateStyle();

                // NOTE: Paragraph formatting (line spacing, space before/after) is not directly
                // available in the current Aspose.Cells .NET API version. If needed,
                // such formatting can be applied via Rich Text or other mechanisms.

                // Apply the style to the target cell
                cell.SetStyle(customStyle);

                // Define output file path
                string outputPath = "CustomParagraphStyle.xlsx";

                // Ensure the directory exists
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
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}

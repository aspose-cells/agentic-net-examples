// Title: How to modify rich‑text in an Excel cell and save the workbook without losing any original formatting using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# change cell rich text while keeping font color and style | preserve cell formatting after updating rich‑text value with Aspose.Cells .NET | save Excel workbook after PutValue without losing existing cell styles in C# | reapply original style to a cell after modifying rich‑text using Aspose.Cells | keep cell borders and number format when changing rich‑text in Aspose.Cells
// Tags: rich‑text cell update Aspose.Cells | retain cell formatting after PutValue Aspose.Cells | reapply saved style Aspose.Cells C# | save Excel workbook preserving formatting Aspose.Cells | Excel rich‑text modification .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, captures the original Style of a cell containing rich‑text, replaces the cell's value with new rich‑text, reapplies the saved Style, and saves the workbook, ensuring that all original formatting attributes such as fonts, colors, borders, and number formats remain unchanged.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Ensure the input file exists; create a simple workbook if it does not.
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found. Creating a sample workbook.");
                    var sampleWorkbook = new Workbook();
                    var sampleSheet = sampleWorkbook.Worksheets[0];
                    sampleSheet.Cells["A1"].PutValue("Sample rich‑text");
                    sampleWorkbook.Save(inputPath);
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Target the cell that contains rich‑text (e.g., A1)
                Cell cell = worksheet.Cells["A1"];

                // Preserve the cell's original style (includes all formatting attributes)
                Style originalStyle = cell.GetStyle();

                // Modify the cell's rich‑text value while keeping the original style intact
                cell.PutValue("New rich‑text content");

                // Re‑apply the original style to the cell
                cell.SetStyle(originalStyle);

                // Save the workbook – all original formatting attributes are retained
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}

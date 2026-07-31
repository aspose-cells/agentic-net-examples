// Title: Aspose.Cells C# – Apply Double‑Line Outline Borders to a Summary Block Range
// Description: Creates a workbook, defines a range (e.g., B2:D5) as a summary block, and uses SetOutlineBorders with CellBorderType.Double and a black color to draw double‑line borders on the range's outer edges before saving the file.
// Keywords: Aspose.Cells SetOutlineBorders | double line border C# | outline borders range | summary block formatting Aspose | C# Excel double border | Aspose.Cells workbook styling
// Common Searches: Aspose.Cells double outline border example | SetOutlineBorders outer edges C# | how to add double line borders to a range in Aspose.Cells | C# apply double border to summary block Excel | Aspose.Cells border style double line
// Developer Intent: Add a double‑line border around the outer edges of a specified cell range.
// Use Cases: Design a report header separated from data with a double‑line outline. | Highlight totals or summary sections in financial spreadsheets. | Visually distinguish a data summary block in exported Excel files.
// AI Prompts: Generate C# code with Aspose.Cells that applies a double‑line outline border to a user‑defined range. | Show how to customize outline border color and style for multiple ranges in a workbook using Aspose.Cells. | Explain combining SetOutlineBorders with conditional formatting to auto‑highlight summary blocks based on cell values.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Example that creates a workbook and outlines a summary block with double borders.
    // Creates a workbook, defines a range (e.g., B2:D5) as a summary block, and uses SetOutlineBorders with CellBorderType.Double and a black color to draw double‑line borders on the range's outer edges before saving the file.
    public class SummaryBlockOutline
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Define the range that represents the summary block (adjust as needed).
                Aspose.Cells.Range summaryRange = worksheet.Cells.CreateRange("B2:D5");

                // Apply a double line style border to all outer edges of the range using black color.
                summaryRange.SetOutlineBorders(CellBorderType.Double, Color.Black);

                // Save the workbook with the applied outline borders.
                workbook.Save("SummaryBlockOutline.xlsx");
                Console.WriteLine("Workbook saved as SummaryBlockOutline.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application.
    internal class Program
    {
        private static void Main(string[] args)
        {
            SummaryBlockOutline.Run();
        }
    }
}

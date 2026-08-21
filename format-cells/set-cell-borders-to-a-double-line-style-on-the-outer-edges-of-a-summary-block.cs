// Title: Aspose.Cells for .NET: Apply Double‑Line Outline Borders to a Summary Block (C#)
// Description: Creates a new workbook, defines a range (e.g., B2:D5) as a summary block, and uses the SetOutlineBorders method with CellBorderType.Double and a black color to draw a double‑line outline around the range before saving the file.
// Keywords: Aspose.Cells C# | .NET Excel formatting | SetOutlineBorders | CellBorderType.Double | double line border | outline border range | summary block styling | Excel range border | B2:D5 border example | format cells Aspose
// Common Searches: Aspose.Cells set double line outline border C# | How to add outer borders to a range with Aspose.Cells | SetOutlineBorders example .NET | Apply double border to summary block Excel | C# code for double line border in Aspose.Cells
// Developer Intent: Add a black double‑line outline to the outer edges of a specified cell range in an Excel worksheet using Aspose.Cells for .NET.
// Use Cases: Design a report header that stands out with a double‑line frame. | Highlight a totals or summary section in financial spreadsheets. | Separate a data‑summary area from the main sheet for clearer readability.
// AI Prompts: Generate C# code with Aspose.Cells that adds a double‑line outline to a dynamic range based on the size of the data. | Show how to set different colors or styles for each side of an outline border using Aspose.Cells. | Provide an example that applies double‑line outline borders to multiple non‑contiguous ranges in the same worksheet.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a new workbook, defines a range (e.g., B2:D5) as a summary block, and uses the SetOutlineBorders method with CellBorderType.Double and a black color to draw a double‑line outline around the range before saving the file.
    public class SummaryBlockOutline
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Define the range that represents the summary block (adjust as needed)
                Aspose.Cells.Range summaryRange = worksheet.Cells.CreateRange("B2:D5");

                // Apply a double line style border to all outer edges of the range
                summaryRange.SetOutlineBorders(CellBorderType.Double, Color.Black);

                // Save the workbook
                string outputPath = "SummaryBlockOutline.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            SummaryBlockOutline.Run();
        }
    }
}

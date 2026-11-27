// Title: How to use Aspose.Cells for .NET to conditionally format cells with QuotePrefix disabled and apply a background color
// Common Searches: Aspose.Cells C# conditional formatting cells where QuotePrefix is false | apply background color to Excel cells without text prefix using Aspose.Cells .NET | detect and highlight cells lacking QuotePrefix flag in a workbook with Aspose.Cells | C# iterate over worksheet cells and set shading based on QuotePrefix property | Aspose.Cells style flag CellShading example for conditional formatting
// Tags: Aspose.Cells style flag CellShading | Excel QuotePrefix detection with Aspose.Cells | conditional background shading for non‑prefixed cells | C# workbook cell style manipulation Aspose.Cells | apply light yellow shading via Aspose.Cells

using System;
using Aspose.Cells;
using System.Drawing;

namespace QuotePrefixConditionalFormatting
{
    // The example creates a new workbook, sets the QuotePrefix flag on specific cells (A1 and B1) to treat their values as text, then loops through all used cells. For each cell whose Style.QuotePrefix property is false, it merges a light‑yellow background style using a StyleFlag that only modifies CellShading, preserving existing formatting. Finally, the workbook is saved as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // ------------------------------------------------------------
                // Populate sample data
                // ------------------------------------------------------------
                // Cell A1 – value that should be treated as text (QuotePrefix = true)
                Cell cellA1 = worksheet.Cells["A1"];
                cellA1.PutValue("12345");
                Style styleA1 = workbook.CreateStyle();
                styleA1.QuotePrefix = true; // Enable QuotePrefix
                // Apply the style (only QuotePrefix flag)
                cellA1.SetStyle(styleA1, new StyleFlag { QuotePrefix = true });

                // Cell A2 – same value but without QuotePrefix (default false)
                worksheet.Cells["A2"].PutValue("12345");

                // Cell B1 – value with leading single quote (Excel treats it as text)
                Cell cellB1 = worksheet.Cells["B1"];
                cellB1.PutValue("'ABC"); // The leading quote is part of the string
                Style styleB1 = workbook.CreateStyle();
                styleB1.QuotePrefix = true;
                cellB1.SetStyle(styleB1, new StyleFlag { QuotePrefix = true });

                // Cell B2 – plain text without QuotePrefix
                worksheet.Cells["B2"].PutValue("ABC");

                // ------------------------------------------------------------
                // Apply conditional formatting: highlight cells where QuotePrefix is false
                // ------------------------------------------------------------
                // Define the highlight style (light yellow background)
                Style highlightStyle = workbook.CreateStyle();
                highlightStyle.ForegroundColor = Color.LightYellow;
                highlightStyle.Pattern = BackgroundType.Solid;

                // Iterate through all used cells
                foreach (Cell cell in worksheet.Cells)
                {
                    // Retrieve the current style of the cell
                    Style currentStyle = cell.GetStyle();

                    // Check the QuotePrefix flag
                    if (!currentStyle.QuotePrefix)
                    {
                        // Merge the highlight background with the existing style
                        currentStyle.ForegroundColor = highlightStyle.ForegroundColor;
                        currentStyle.Pattern = highlightStyle.Pattern;

                        // Apply only the CellShading flag to avoid altering other formatting
                        StyleFlag flag = new StyleFlag { CellShading = true };
                        cell.SetStyle(currentStyle, flag);
                    }
                }

                // ------------------------------------------------------------
                // Save the workbook
                // ------------------------------------------------------------
                string outputPath = "QuotePrefixConditionalFormatting.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}

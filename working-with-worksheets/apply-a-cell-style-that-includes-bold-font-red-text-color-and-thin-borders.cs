// Title: Apply Bold Red Font and Thin Black Borders to a Cell Using Aspose.Cells for .NET (C#)
// Description: This C# example demonstrates how to create a Workbook with Aspose.Cells, define a Style that makes the font bold and red, adds thin black borders on all sides, applies the style to cell B2, writes text, and saves the workbook as StyledCell.xlsx.
// Keywords: Aspose.Cells C# style | bold red font Excel | thin borders Aspose.Cells | format cell programmatically .NET | Excel cell styling Aspose | C# Excel border color | Aspose.Cells workbook example
// Common Searches: Aspose.Cells set bold red text with borders | C# apply thin border to Excel cell using Aspose | how to style a cell in Aspose.Cells .NET | change font color and border in Aspose.Cells
// Developer Intent: Format a specific Excel cell with bold red text and thin black borders using Aspose.Cells for .NET.
// Use Cases: Design header rows that stand out with bold red headings and thin borders. | Highlight critical values in financial dashboards by applying a consistent style. | Create pre‑formatted template cells for user input in generated reports.
// AI Prompts: Write C# code that applies a bold, red font and thin black borders to a range of cells with Aspose.Cells. | Show how to create a reusable Style object in Aspose.Cells and assign it to multiple cells efficiently. | Explain how to modify only the border color of an existing bold red style in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This C# example demonstrates how to create a Workbook with Aspose.Cells, define a Style that makes the font bold and red, adds thin black borders on all sides, applies the style to cell B2, writes text, and saves the workbook as StyledCell.xlsx.
    public class ApplyBoldRedThinBorderStyle
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Create a new style object
            Style style = workbook.CreateStyle();

            // Set the font to bold and red
            style.Font.IsBold = true;
            style.Font.Color = Color.Red;

            // Apply thin black borders on all four sides
            style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);

            // Apply the style to a specific cell (B2)
            Cell cell = cells["B2"];
            cell.PutValue("Styled Text");
            cell.SetStyle(style);

            // Save the workbook to a file
            string outputPath = "StyledCell.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}

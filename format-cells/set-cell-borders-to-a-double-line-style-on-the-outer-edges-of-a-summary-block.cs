// Title: Add double‑line outer borders to a specific summary block range in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a Style with CellBorderType.Double for the top, bottom, left, and right borders and applies it to a defined CellArea using Aspose.Cells. | Show how to use a StyleFlag to apply only border settings to a range while preserving other cell formatting in an Aspose.Cells workbook. | Write a C# example that fills cells A1:D5 with sample data, adds double‑line outer borders around that block, and saves the file as an .xlsx document.
// Common Searches: Aspose.Cells C# set double border on range edges only | how to apply outer borders to a block of cells in Aspose.Cells .NET | using StyleFlag to format only borders in Aspose.Cells workbook | create double line border around summary table with Aspose.Cells C# | apply border style to specific CellArea in Aspose.Cells example
// Tags: Aspose.Cells double outer border style C# | StyleFlag border-only formatting Aspose.Cells | apply borders to CellArea Aspose.Cells | Excel double line border Aspose.Cells .NET | summary block range styling Aspose.Cells

using System;
using Aspose.Cells;
using System.Drawing;

// Demonstrates creating a workbook, defining a summary block (A1:D5), filling it with sample data, building a style with double‑line top, bottom, left, and right borders, applying the style only to the range using a StyleFlag, and saving the result as SummaryBlock.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the summary block range (e.g., A1:D5)
            CellArea summaryArea = new CellArea
            {
                StartRow = 0,      // Row 1 (zero‑based)
                StartColumn = 0,   // Column A
                EndRow = 4,        // Row 5
                EndColumn = 3      // Column D
            };

            // (Optional) Fill the block with sample data
            for (int i = summaryArea.StartRow; i <= summaryArea.EndRow; i++)
            {
                for (int j = summaryArea.StartColumn; j <= summaryArea.EndColumn; j++)
                {
                    sheet.Cells[i, j].PutValue($"R{i + 1}C{j + 1}");
                }
            }

            // Create a style that uses a double line for the outer borders
            Style borderStyle = workbook.CreateStyle();

            // Set double line style for each outer border
            borderStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Double;
            borderStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Double;
            borderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Double;
            borderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Double;

            // Apply the border style only to the borders of the range
            Aspose.Cells.Range summaryRange = sheet.Cells.CreateRange(
                summaryArea.StartRow,
                summaryArea.StartColumn,
                summaryArea.EndRow - summaryArea.StartRow + 1,
                summaryArea.EndColumn - summaryArea.StartColumn + 1);

            // Use a StyleFlag to apply only the border settings
            StyleFlag flag = new StyleFlag();
            flag.All = false;
            flag.Borders = true;   // Apply only border-related settings

            summaryRange.ApplyStyle(borderStyle, flag);

            // Save the workbook to a file
            workbook.Save("SummaryBlock.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

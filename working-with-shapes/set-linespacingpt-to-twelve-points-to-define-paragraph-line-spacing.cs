// Title: How to set a 12‑point line spacing for wrapped paragraph text in an Excel cell with Aspose.Cells for .NET
// AI Prompts: Write C# code that uses Aspose.Cells to enable text wrapping and apply a 12‑point line spacing to a paragraph inside a specific worksheet cell. | Show how to adjust row height together with text wrapping to achieve exact 12‑point line spacing for cell content using the Aspose.Cells .NET API. | Provide a minimal example that creates a workbook, inserts wrapped text, and configures the cell to display a 12‑point line spacing with Aspose.Cells.
// Common Searches: Aspose.Cells set exact line spacing for wrapped text in a cell | C# Aspose.Cells 12 point paragraph spacing in Excel worksheet | How to control line spacing of cell text using Aspose.Cells .NET | Adjust row height to match paragraph line spacing with Aspose.Cells | Set paragraph line spacing in Excel cell programmatically with Aspose.Cells
// Tags: Aspose.Cells set cell line spacing | Aspose.Cells text wrapping row height | C# Excel paragraph spacing Aspose.Cells | Aspose.Cells line spacing 12pt | Excel cell formatting Aspose.Cells .NET

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, writes a paragraph into cell A1, enables text wrapping, sets the row height to 12 points to approximate a 12‑point line spacing, and saves the workbook as Output.xlsx using Aspose.Cells for .NET.
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

            // Put sample text into cell A1
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("This paragraph uses a line spacing of 12 points.");

            // Enable text wrapping for the cell
            Style style = cell.GetStyle();
            style.IsTextWrapped = true;
            cell.SetStyle(style);

            // Set the row height to 12 points (approximate line spacing)
            sheet.Cells.SetRowHeight(0, 12);

            // Define output file path
            string outputPath = "Output.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

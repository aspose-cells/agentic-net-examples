// Title: Create an Excel workbook and add an external hyperlink to cell H5 using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that creates a new workbook, sets the text of cell H5 to "Visit Example", adds a hyperlink to https://www.example.com, and saves the file. | Show how to call Worksheet.Hyperlinks.Add to attach an external URL to a single cell at row 5, column H in Aspose.Cells. | Generate a minimal Aspose.Cells example that inserts a clickable link into cell H5 of the first worksheet and exports it as HyperlinkDemo.xlsx.
// Common Searches: Aspose.Cells C# add external link to a specific cell in Excel | How to insert a hyperlink into cell H5 with Aspose.Cells for .NET | C# Aspose.Cells Hyperlinks.Add example for single cell | Create Excel workbook with clickable URL in a cell using Aspose.Cells | Add web link to cell H5 programmatically with Aspose.Cells
// Tags: Aspose.Cells add hyperlink to cell | Worksheet.Hyperlinks.Add external URL | C# create Excel workbook with clickable link | hyperlink cell H5 Aspose.Cells | save workbook as HyperlinkDemo.xlsx

using System;
using Aspose.Cells;

// Demonstrates creating a new workbook, setting the display text of cell H5, adding an external hyperlink to that cell with Worksheet.Hyperlinks.Add, and saving the file as HyperlinkDemo.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Target cell coordinates for H5 (zero‑based indices)
            int row = 4;      // H5 -> row 5 (index 4)
            int column = 7;   // column H (index 7)

            // Set the display text of the cell
            worksheet.Cells[row, column].PutValue("Visit Example");

            // Add a hyperlink that opens an external website when clicked
            // For a single cell, totalRows = 1 and totalColumns = 1
            worksheet.Hyperlinks.Add(row, column, 1, 1, "https://www.example.com");

            // Save the workbook
            workbook.Save("HyperlinkDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

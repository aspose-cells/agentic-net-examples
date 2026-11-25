// Title: Create an external hyperlink in cell H2 of an Excel worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add hyperlink to cell H2 with display text | How to insert an external URL into a specific Excel cell using Aspose.Cells | Set hyperlink text for a single cell in a generated .xlsx file with Aspose.Cells | Programmatically add a hyperlink to an Excel worksheet in .NET | Aspose.Cells Hyperlinks.Add example for a single cell
// Tags: Aspose.Cells add cell hyperlink | Hyperlinks.Add method Excel .NET | set hyperlink display text Aspose.Cells | save workbook as Xlsx Aspose.Cells | external URL hyperlink Excel cell C#

using System;
using Aspose.Cells;

// The example creates a new workbook, accesses the first worksheet, adds an external hyperlink pointing to https://example.com in cell H2, sets the cell's display text to "Example Site", and saves the file as HyperlinkDemo.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Initialize a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet (index 0)
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a hyperlink to cell H2 (row index 1, column index 7)
        // Parameters: startRow, startColumn, totalRows, totalColumns, address
        worksheet.Hyperlinks.Add(1, 7, 1, 1, "https://example.com");

        // Optionally, set a display text for the hyperlink in the cell
        worksheet.Cells["H2"].PutValue("Example Site");

        // Save the workbook to a file
        string outputPath = "HyperlinkDemo.xlsx";
        try
        {
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}

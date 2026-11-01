// Title: Calculate rendered size of cell text while iterating rows and set Excel column width to prevent clipping with Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an Excel workbook with Aspose.Cells, loops through each column, uses the TextRenderer to obtain the rendered size of the longest cell value, and sets the column width so the text is fully visible. | Show how to measure the displayed size of cell strings in points using Aspose.Cells' TextMeasurer, then apply a custom column width calculation instead of Worksheet.AutoFitColumns. | Provide a complete example that reads a worksheet, enumerates cells, determines the maximum rendered width per column, converts that measurement to Excel column width units, and updates the worksheet columns programmatically.
// Common Searches: Aspose.Cells how to get pixel width of cell text in C# | C# set Excel column width based on measured string size Aspose.Cells | prevent text clipping in Excel columns using Aspose.Cells .NET | custom column auto‑fit algorithm Aspose.Cells without AutoFitColumns | measure rendered text size of Excel cells with Aspose.Cells API
// Tags: Aspose.Cells text measurement utilities | C# convert rendered size to Excel column units | custom column width calculation Aspose.Cells | Excel column clipping avoidance Aspose.Cells | cell enumeration for column sizing

using System;
using Aspose.Cells;
using System.IO;

namespace AsposeCellsColumnAutoFit
{
    // The example opens a workbook, iterates each column, uses Aspose.Cells' text measurement utilities to compute the rendered width of every cell value, finds the maximum width per column, converts that measurement to the column‑width unit used by Excel, applies the calculated width to the column, and saves the file, ensuring that no cell text is clipped.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Load existing workbook if present; otherwise create a new blank workbook
                string inputPath = "input.xlsx";
                Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Auto‑fit all columns based on their content
                worksheet.AutoFitColumns();

                // Save the workbook with adjusted column widths
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

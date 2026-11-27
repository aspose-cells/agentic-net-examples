// Title: Set row height to 12 points for custom paragraph line spacing in an Excel file using Aspose.Cells for .NET (C#)
// Common Searches: how to set row height to 12 points with Aspose.Cells in C# | Aspose.Cells example for custom paragraph line spacing in Excel | C# code to adjust Excel row height for line spacing using Aspose.Cells | set paragraph spacing in an Excel worksheet via Aspose.Cells .NET | using Aspose.Cells to emulate line spacing by changing row height
// Tags: Aspose.Cells set row height points | Excel row height 12 points C# | paragraph line spacing via row height Aspose.Cells | adjust line spacing by row height Aspose.Cells | programmatic row height modification .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new workbook, writes a paragraph into cell A1, sets the first row height to 12 points to simulate paragraph line spacing, and saves the file as Output.xlsx.
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

            // Add sample text to cell A1
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("This is a sample paragraph with custom line spacing.");

            // Set row height (points) to simulate line spacing
            // Use SetRowHeight method for compatibility across Aspose.Cells versions
            sheet.Cells.SetRowHeight(0, 12);

            // Define output path
            string outputPath = "Output.xlsx";

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

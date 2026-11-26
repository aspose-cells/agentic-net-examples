// Title: Set row height to control line spacing and emulate character spacing for balanced text in Aspose.Cells for .NET
// Common Searches: how to change row height for line spacing in Aspose.Cells C# | simulate character spacing in Excel cells using Aspose.Cells .NET | balanced text appearance in spreadsheet with Aspose.Cells | adjust vertical spacing of text in Aspose.Cells workbook | increase row height to improve text readability Aspose.Cells C#
// Tags: line spacing via row height Aspose.Cells | character spacing workaround Excel .NET | balanced text formatting Aspose.Cells | style API constraints Aspose.Cells | vertical spacing tuning C# spreadsheet

using Aspose.Cells;
using System;
using System.IO;

// The example creates a workbook, writes text to cell A1, and demonstrates how to adjust the first row's height to influence line spacing. Because Aspose.Cells does not expose a direct character spacing property, the code shows a workaround to emulate character spacing by modifying row height, then saves the file as AdjustedSpacing.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Write sample text into a cell
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("Balanced text appearance");

            // Retrieve the cell's style
            Style style = cell.GetStyle();

            // Aspose.Cells does not expose a direct character spacing property.
            // If needed, other formatting options can be applied here.

            // Apply the (unchanged) style back to the cell
            cell.SetStyle(style);

            // Adjust line spacing by setting the row height.
            Row row = sheet.Cells.Rows[0];
            row.Height = 25; // set a height that balances with the character spacing

            // Ensure the output directory exists
            string outputPath = "AdjustedSpacing.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

// Title: Set a cell’s font size to 12 points with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# change font size of cell A1 to 12 points | How to apply 12pt font to specific text in an Excel cell using Aspose.Cells | C# example for setting cell text font size with Aspose.Cells library | Update Excel cell style to 12 point font in .NET using Aspose | Set whole cell font size when RichTextPortion editing is unavailable in Aspose.Cells
// Tags: Aspose.Cells set cell font size .NET | C# Excel cell style modification Aspose | apply 12pt font to cell text Aspose.Cells | modify workbook cell formatting using Aspose.Cells C# | save updated Excel file with new font size Aspose

using Aspose.Cells;
using System;
using System.IO;

// The sample loads (or creates) an Excel workbook, accesses cell A1 on the first worksheet, retrieves the cell's style, sets the font size to 12 points (applied to the entire cell because RichTextPortion editing is not supported in this version), reapplies the style, and saves the workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists; create a simple workbook if it does not.
            if (!File.Exists(inputPath))
            {
                var tempWb = new Workbook();
                var tempWs = tempWb.Worksheets[0];
                tempWs.Cells["A1"].PutValue("Sample Text");
                tempWb.Save(inputPath);
            }

            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet.
            Worksheet worksheet = workbook.Worksheets[0];

            // Target cell (e.g., A1).
            Cell cell = worksheet.Cells["A1"];

            // Retrieve the cell's current style.
            Style style = cell.GetStyle();

            // Set the font size for the entire cell (as RichTextPortion manipulation is not available in this version).
            style.Font.Size = 12;

            // Apply the updated style back to the cell.
            cell.SetStyle(style);

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// Title: Highlight cells containing the word “Pending” with a yellow fill in all worksheets using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an existing .xlsx file with Aspose.Cells, searches every worksheet for cells whose text includes “Pending” (case‑insensitive), and applies a solid yellow background style to each matching cell. | Create a reusable C# method that takes a workbook path, a search keyword, and a Color, then uses Aspose.Cells to locate all cells containing the keyword and set their fill pattern to solid with the specified color. | Extend the sample to process multiple Excel files, allowing the keyword and highlight color to be passed as parameters, and apply the formatting with Aspose.Cells.
// Common Searches: Aspose.Cells C# find and color cells that contain a specific substring | How to apply yellow background to cells with text 'Pending' in an Excel workbook using Aspose.Cells | Search all worksheets for a keyword and set cell style with Aspose.Cells .NET | C# Aspose.Cells iterate through cells and change fill color based on string value
// Tags: search cells by string Aspose.Cells | apply solid fill style Aspose.Cells C# | highlight matching cells Excel .NET | iterate worksheets Aspose.Cells | case‑insensitive text search Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// Loads an existing .xlsx workbook, scans every worksheet for cells whose string value includes the word “Pending” (ignoring case), sets a solid yellow background on those cells using Aspose.Cells styling, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the used range of the sheet
                var cells = sheet.Cells;
                var maxRow = cells.MaxDataRow;
                var maxColumn = cells.MaxDataColumn;

                // Scan each cell for the text "Pending"
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxColumn; col++)
                    {
                        Cell cell = cells[row, col];
                        if (cell.Type == CellValueType.IsString && 
                            cell.StringValue.IndexOf("Pending", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // Highlight the cell with a yellow fill
                            Style style = cell.GetStyle();
                            style.ForegroundColor = Color.Yellow;
                            style.Pattern = BackgroundType.Solid;
                            cell.SetStyle(style);
                        }
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

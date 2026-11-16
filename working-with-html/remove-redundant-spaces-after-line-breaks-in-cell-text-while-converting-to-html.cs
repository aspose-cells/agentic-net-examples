// Title: Trim extra spaces after line breaks in Excel cell strings before saving as HTML with Aspose.Cells for .NET
// Common Searches: c# Aspose.Cells remove spaces that follow a newline character in worksheet cells | how to clean up line‑break whitespace in Excel cells prior to HTML export using Aspose.Cells | replace \n followed by spaces in Excel cell values with a single line break in .NET | regex to collapse newline‑space sequences in Aspose.Cells before generating HTML
// Tags: regex newline space removal Aspose.Cells | cell string cleanup before HTML export | Aspose.Cells update cell value C# | Excel to HTML whitespace normalization | line break whitespace handling .NET

using System;
using Aspose.Cells;

// The example loads an Excel workbook, iterates through each used cell, and for string cells applies a regular expression that replaces a line‑break character followed by one or more spaces with a single line break. Modified cells are updated, and the workbook is saved as an HTML file.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the used range of the worksheet
            var cells = sheet.Cells;
            var maxRow = cells.MaxDataRow;
            var maxCol = cells.MaxDataColumn;

            // Loop through each cell in the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];

                    // Process only cells that contain string values
                    if (cell.Type == CellValueType.IsString && cell.Value != null)
                    {
                        string text = cell.StringValue;

                        // Replace line break followed by one or more spaces with a single line break
                        // Excel line break is represented by '\n' (char 10)
                        string cleaned = System.Text.RegularExpressions.Regex.Replace(
                            text,
                            @"\n\s+",
                            "\n"
                        );

                        // Update the cell only if changes were made
                        if (!object.ReferenceEquals(text, cleaned))
                        {
                            cell.PutValue(cleaned);
                        }
                    }
                }
            }
        }

        // Save the workbook as HTML (replace with your desired output path)
        workbook.Save("output.html", SaveFormat.Html);
    }
}

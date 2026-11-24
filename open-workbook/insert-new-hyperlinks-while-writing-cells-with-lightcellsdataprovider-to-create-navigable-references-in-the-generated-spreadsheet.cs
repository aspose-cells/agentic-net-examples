// Title: How to add clickable hyperlinks with screen tips to Excel cells while populating a worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add hyperlink to a specific cell with custom screen tip | programmatically insert URL links into an Excel worksheet using Aspose.Cells | Aspose.Cells example for writing cells and attaching hyperlinks in .NET | create clickable Excel cells with Aspose.Cells and set tooltip text | populate Excel sheet with data and add hyperlinks using Aspose.Cells C#
// Tags: excel hyperlink creation Aspose.Cells | custom screen tip for hyperlinks Aspose.Cells | populate worksheet with URLs Aspose.Cells | save workbook as xlsx Aspose.Cells | hyperlink column handling Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The sample creates a new workbook, writes a header row and data rows containing display text and URLs, adds a hyperlink to each URL cell (excluding the header), assigns a custom screen tip based on the adjacent name, and saves the result as Hyperlinks.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Sample data: first column is display text, second column is the hyperlink URL
            var data = new List<string[]>
            {
                new[] { "Name", "URL" },
                new[] { "Google",   "https://www.google.com" },
                new[] { "Microsoft","https://www.microsoft.com" },
                new[] { "OpenAI",   "https://www.openai.com" }
            };

            // Populate cells and add hyperlinks for the URL column (except header row)
            for (int row = 0; row < data.Count; row++)
            {
                for (int col = 0; col < data[row].Length; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    cell.PutValue(data[row][col]);

                    // Add hyperlink to the URL column (skip header)
                    if (col == 1 && row > 0)
                    {
                        // Add a hyperlink to the single cell; Add returns the index of the hyperlink
                        int linkIndex = sheet.Hyperlinks.Add(row, col, 1, 1, data[row][col]);
                        Hyperlink link = sheet.Hyperlinks[linkIndex];
                        link.ScreenTip = "Open " + data[row][0];
                    }
                }
            }

            // Save the generated workbook
            string outputPath = "Hyperlinks.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}

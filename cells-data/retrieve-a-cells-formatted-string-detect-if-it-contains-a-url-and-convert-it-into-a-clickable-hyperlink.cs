// Title: Read a cell's displayed text, detect embedded URLs, and convert them into clickable hyperlinks with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells that reads the formatted string from a worksheet cell, uses a regular expression to locate an http or https URL, and inserts a hyperlink into that same cell. | Demonstrate how to apply the conventional hyperlink appearance (blue font color and single underline) to a cell after adding a hyperlink with Aspose.Cells.
// Common Searches: Aspose.Cells C# convert URL found in cell text to clickable hyperlink | detect and replace URLs with hyperlinks in Excel using Aspose.Cells .NET | add hyperlink to Excel cell based on regex match with Aspose.Cells | C# Aspose.Cells create hyperlink from cell string value and style it | read formatted cell value and insert hyperlink programmatically in Aspose.Cells
// Tags: read cell display text Aspose.Cells | regex URL detection worksheet cell C# | insert cell hyperlink Aspose.Cells | apply hyperlink style blue underline Aspose.Cells | save workbook after hyperlink insertion Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;

// The example loads or creates an Excel workbook, reads the formatted text of a target cell, uses a regular expression to find an http/https URL, adds a hyperlink that covers the same cell, applies the standard blue underlined hyperlink style, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets.Add();
            }

            Worksheet worksheet = workbook.Worksheets[0];

            // Target cell (modify address as needed)
            Cell cell = worksheet.Cells["A1"];

            // Retrieve the cell's formatted string
            string formattedText = cell.StringValue;

            // Detect a URL in the formatted text
            Regex urlRegex = new Regex(@"\bhttps?://[^\s]+", RegexOptions.IgnoreCase);
            Match match = urlRegex.Match(formattedText);

            if (match.Success)
            {
                string url = match.Value;

                // Add a hyperlink that covers the same cell
                worksheet.Hyperlinks.Add(cell.Row, cell.Column, 1, 1, url);

                // Apply typical hyperlink style (blue & underlined)
                Style hyperlinkStyle = workbook.CreateStyle();
                hyperlinkStyle.Font.Color = Color.Blue;
                hyperlinkStyle.Font.Underline = FontUnderlineType.Single;
                cell.SetStyle(hyperlinkStyle);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

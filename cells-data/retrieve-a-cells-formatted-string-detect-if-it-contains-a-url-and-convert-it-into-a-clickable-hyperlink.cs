// Title: Read a cell's displayed text, detect an embedded URL, and add a clickable hyperlink with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# read cell displayed value and convert detected URL to hyperlink | C# regex to find HTTP link in Excel cell and create hyperlink using Aspose.Cells | Add hyperlink to Excel cell based on its string content with Aspose.Cells .NET | Programmatically turn plain URL text in a worksheet cell into a clickable link using Aspose.Cells | Save workbook after inserting hyperlink into cell A1 with Aspose.Cells C#
// Tags: read cell displayed text Aspose.Cells | detect url in cell using regex C# | add hyperlink to worksheet cell Aspose.Cells | hyperlink insertion based on cell content .NET | save workbook after hyperlink Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;
using System.Text.RegularExpressions;

// The example loads an existing workbook (or creates a new one), reads the displayed text of a specified cell, uses a regular expression to locate an HTTP/HTTPS URL, adds a hyperlink that covers the same cell when a URL is found, ensures the output directory exists, and saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook or create a new one if the file is missing
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Ensure there is at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                workbook.Worksheets.Add();
            }

            Worksheet sheet = workbook.Worksheets[0]; // first worksheet

            // Specify the cell to process (change as needed)
            Cell cell = sheet.Cells["A1"];

            // Retrieve the cell's displayed text
            string formattedText = cell.StringValue;

            // Detect a URL inside the formatted text using a simple regex
            Regex urlRegex = new Regex(@"\bhttps?://[^\s]+", RegexOptions.IgnoreCase);
            Match match = urlRegex.Match(formattedText);

            if (match.Success)
            {
                string url = match.Value;

                // Add a hyperlink covering the same cell (1 row x 1 column)
                sheet.Hyperlinks.Add(cell.Row, cell.Column, 1, 1, url);
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}

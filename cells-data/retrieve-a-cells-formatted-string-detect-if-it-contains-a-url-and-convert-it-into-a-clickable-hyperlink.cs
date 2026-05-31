using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data: some cells contain URLs, others do not
            sheet.Cells["A1"].PutValue("Visit https://www.example.com for more info");
            sheet.Cells["A2"].PutValue("Just some text without a link");
            sheet.Cells["A3"].PutValue("http://test.com");

            // Iterate through all used cells
            AsposeRange usedRange = sheet.Cells.MaxDisplayRange;
            foreach (Cell cell in usedRange)
            {
                // Retrieve the formatted (display) string of the cell
                string displayText = cell.DisplayStringValue;

                // Detect a URL using a simple regular expression
                if (Regex.IsMatch(displayText, @"https?://\S+"))
                {
                    // Extract the first URL found
                    string url = Regex.Match(displayText, @"https?://\S+").Value;

                    // If the cell does not already contain a hyperlink, add one
                    if (!cell.ContainsExternalLink)
                    {
                        // Add a hyperlink to the cell (using the rule Hyperlinks.Add(string, int, int, string))
                        int hyperlinkIndex = sheet.Hyperlinks.Add(cell.Name, 1, 1, url);

                        // Set the displayed text of the hyperlink to the original cell text
                        sheet.Hyperlinks[hyperlinkIndex].TextToDisplay = displayText;
                    }
                }
            }

            // Ensure output directory exists
            string outputPath = "Output.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
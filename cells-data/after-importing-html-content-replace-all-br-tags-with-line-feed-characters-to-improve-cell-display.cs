// Title: Replace HTML <br> tags with newline characters in Excel cells after loading HTML with Aspose.Cells for .NET
// Common Searches: how to convert <br> tags to newline in cells after loading HTML with Aspose.Cells C# | Aspose.Cells replace HTML line break tags in worksheet string values | C# remove <br> elements from imported HTML using HtmlLoadOptions | post‑import HTML cleanup Aspose.Cells replace br with \n | regex replace <br> in Excel cells using Aspose.Cells .NET
// Tags: replace br tags Aspose.Cells | html line break conversion .NET workbook | post load cell string cleanup | regex br to newline Excel cells | HtmlLoadOptions cell value processing

using Aspose.Cells;
using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

// The example reads an HTML file, loads it into an Aspose.Cells Workbook via HtmlLoadOptions, iterates through all worksheets and string cells, and uses a case‑insensitive regular expression to replace <br>, <br/>, and <br /> tags with a line feed character. Updated cell values are saved to an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the HTML file to import
            string htmlPath = "input.html";

            // Verify that the input file exists
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Input file not found: {htmlPath}");
                return;
            }

            // Read HTML content from the file
            string htmlContent = File.ReadAllText(htmlPath);

            // Load HTML content into a workbook using HtmlLoadOptions and a memory stream
            Workbook workbook;
            using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(htmlContent)))
            {
                HtmlLoadOptions loadOptions = new HtmlLoadOptions();
                workbook = new Workbook(ms, loadOptions);
            }

            // Iterate through all worksheets and replace <br> tags with line feeds in string cells
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];
                        if (cell.Type == CellValueType.IsString)
                        {
                            string text = cell.StringValue;
                            if (!string.IsNullOrEmpty(text) && text.IndexOf("<br", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                // Replace <br>, <br/>, <br /> (case‑insensitive) with "\n"
                                string newText = Regex.Replace(text, @"<br\s*/?>", "\n", RegexOptions.IgnoreCase);
                                cell.PutValue(newText);
                            }
                        }
                    }
                }
            }

            // Save the modified workbook to an Excel file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

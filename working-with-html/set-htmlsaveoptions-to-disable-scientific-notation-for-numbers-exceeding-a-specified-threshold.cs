// Title: Disable scientific notation for large numeric values when saving an Excel workbook to HTML with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export to HTML without scientific notation for numbers larger than 1e10 | force plain‑text representation of big numbers in HTML output using Aspose.Cells | C# convert numeric cells to string before saving workbook as HTML with Aspose.Cells | Set HtmlSaveOptions to keep large numbers from showing exponential format in Aspose.Cells | prevent exponential notation in HTML files generated from Excel using Aspose.Cells .NET
// Tags: Aspose.Cells HTML export numeric formatting | disable scientific notation Aspose.Cells | C# preprocess cells for HTML conversion | HtmlSaveOptions custom number representation | large numbers plain text Aspose.Cells HTML

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, scans all cells, and replaces any numeric value whose absolute magnitude meets or exceeds a defined threshold with a plain‑text string using ToString("0"). It then configures HtmlSaveOptions (e.g., ExportGridLines) and saves the workbook as an HTML file, ensuring large numbers are displayed without scientific notation.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";
        const string outputFile = "output.html";
        const double scientificNotationThreshold = 1e10;

        try
        {
            // Ensure the input file exists; create a simple workbook if it does not.
            if (!File.Exists(inputFile))
            {
                var tempWb = new Workbook();
                var tempSheet = tempWb.Worksheets[0];
                tempSheet.Cells["A1"].PutValue(1234567890123); // Sample large number
                tempWb.Save(inputFile);
                Console.WriteLine($"Sample input file \"{inputFile}\" created.");
            }

            // Load the workbook.
            var workbook = new Workbook(inputFile);

            // Pre‑process cells that exceed the threshold to force plain‑text representation.
            var sheet = workbook.Worksheets[0];
            var cells = sheet.Cells;
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    var cell = cells[row, col];
                    if (cell.Type == CellValueType.IsNumeric)
                    {
                        double value = cell.DoubleValue;
                        if (Math.Abs(value) >= scientificNotationThreshold)
                        {
                            // Convert to a non‑scientific string.
                            cell.PutValue(value.ToString("0"));
                        }
                    }
                }
            }

            // Configure HTML save options (no scientific notation conversion property needed).
            var htmlOptions = new HtmlSaveOptions
            {
                // Example: keep grid lines visible.
                ExportGridLines = true
            };

            // Save as HTML.
            workbook.Save(outputFile, htmlOptions);
            Console.WriteLine($"HTML file saved successfully to \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

// Title: How to center an Excel chart on a PDF page using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export chart to PDF with horizontal and vertical centering | center chart in PDF output when converting an Excel workbook with Aspose.Cells | remove page margins to make a chart fill the PDF page using Aspose.Cells .NET | set print area for a specific chart before saving as PDF in C# Aspose.Cells | adjust worksheet PageSetup to center chart during PDF conversion with Aspose.Cells
// Tags: chart centering PDF export Aspose.Cells | worksheet pagesetup margins removal Aspose.Cells | print area configuration chart Aspose.Cells | C# Aspose.Cells PDF layout adjustment | Excel chart to PDF centered Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an Excel workbook, checks for a chart, configures the worksheet's PageSetup to center content horizontally and vertically, optionally removes all margins, and saves the result as a PDF, ensuring the chart appears centered on the page.
class Program
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("Error: No charts found in the first worksheet.");
                return;
            }

            // Access the first chart
            Chart chart = sheet.Charts[0];

            // OPTIONAL: Define a print area that includes the chart.
            // If the Chart class does not expose UpperLeft/LowerRight properties in the current version,
            // we fall back to printing the entire sheet.
            // Uncomment and adjust the following lines if your version supports these properties.
            /*
            int topRow = chart.UpperLeftRow;
            int leftColumn = chart.UpperLeftColumn;
            int bottomRow = chart.LowerRightRow;
            int rightColumn = chart.LowerRightColumn;

            string startCell = CellsHelper.CellIndexToName(topRow, leftColumn);
            string endCell = CellsHelper.CellIndexToName(bottomRow, rightColumn);
            sheet.PageSetup.PrintArea = $"{startCell}:{endCell}";
            */

            // Center the printed content horizontally and vertically on the PDF page
            sheet.PageSetup.CenterHorizontally = true;
            sheet.PageSetup.CenterVertically = true;

            // Optional: remove margins to let the chart occupy the full page width/height
            sheet.PageSetup.LeftMargin = 0;
            sheet.PageSetup.RightMargin = 0;
            sheet.PageSetup.TopMargin = 0;
            sheet.PageSetup.BottomMargin = 0;

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a PDF file
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

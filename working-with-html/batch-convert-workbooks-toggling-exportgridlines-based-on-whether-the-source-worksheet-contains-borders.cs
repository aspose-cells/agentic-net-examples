// Title: Batch convert Excel workbooks to HTML with conditional gridlines using Aspose.Cells for .NET
// Description: A C# console app that scans every worksheet in each .xlsx file of a folder for cell borders. It saves the workbook as HTML, enabling HtmlSaveOptions.ExportGridLines only when no borders are detected, and processes all files in one run.
// Keywords: Aspose.Cells HTML export | ExportGridLines conditional | detect cell borders C# | batch Excel to HTML | Aspose.Cells .NET example | gridlines toggle based on borders | C# Excel automation | global .NET developers | USA .NET community | GitHub Aspose.Cells sample
// Common Searches: Aspose.Cells export Excel to HTML with gridlines only when no borders | C# batch convert .xlsx files to HTML and disable gridlines if borders exist | how to detect cell borders before saving as HTML using Aspose.Cells | set ExportGridLines dynamically for each workbook in .NET | Aspose.Cells example for conditional HTML save options
// Developer Intent: Automatically convert a directory of Excel workbooks to HTML, turning on gridlines only for workbooks that have no cell borders.
// Use Cases: Generate web‑ready reports from a library of Excel templates while preserving original styling. | Build a server‑side service that receives user spreadsheets, converts them to HTML, and avoids double borders by disabling gridlines when borders are present. | Create an automated publishing pipeline that processes large batches of workbooks and applies visual‑consistency rules without manual intervention.
// AI Prompts: Write a C# function that returns true if any cell in a Workbook has a top, bottom, left, or right border. | Provide an Aspose.Cells .NET snippet that batch converts all .xlsx files in a folder to HTML, setting ExportGridLines based on border detection. | Explain how to extend the code to also toggle ExportColumnHeaders when the first worksheet row contains header text.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// A C# console app that scans every worksheet in each .xlsx file of a folder for cell borders. It saves the workbook as HTML, enabling HtmlSaveOptions.ExportGridLines only when no borders are detected, and processes all files in one run.
class BatchConvertWithGridlineToggle
{
    static void Main()
    {
        // Input and output directories
        string inputFolder = @"C:\InputWorkbooks";
        string outputFolder = @"C:\OutputHtml";

        // Ensure output folder exists
        Directory.CreateDirectory(outputFolder);

        // Process each Excel file in the input folder
        foreach (string sourcePath in Directory.GetFiles(inputFolder, "*.xlsx"))
        {
            // Load the workbook (lifecycle create/load rule)
            Workbook workbook = new Workbook(sourcePath);

            bool hasBorders = false;

            // Scan all worksheets for any cell border
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // Determine the used range to limit the scan
                int maxRow = ws.Cells.MaxDataRow;
                int maxCol = ws.Cells.MaxDataColumn;

                for (int row = 0; row <= maxRow && !hasBorders; row++)
                {
                    for (int col = 0; col <= maxCol && !hasBorders; col++)
                    {
                        // Get the style of the current cell
                        Style style = ws.Cells[row, col].GetStyle();

                        // Check each border side for a non‑None line style
                        if (style.Borders[BorderType.TopBorder].LineStyle != CellBorderType.None ||
                            style.Borders[BorderType.BottomBorder].LineStyle != CellBorderType.None ||
                            style.Borders[BorderType.LeftBorder].LineStyle != CellBorderType.None ||
                            style.Borders[BorderType.RightBorder].LineStyle != CellBorderType.None)
                        {
                            hasBorders = true;
                        }
                    }
                }

                if (hasBorders) break;
            }

            // Prepare HTML save options and toggle ExportGridLines
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Export gridlines only when the source does NOT contain borders
                ExportGridLines = !hasBorders,
                ExportActiveWorksheetOnly = false
            };

            // Build output file path (same name with .html extension)
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(sourcePath);
            string destPath = Path.Combine(outputFolder, fileNameWithoutExt + ".html");

            // Save the workbook as HTML with the configured options (lifecycle save rule)
            workbook.Save(destPath, htmlOptions);
        }

        Console.WriteLine("Batch conversion completed.");
    }
}

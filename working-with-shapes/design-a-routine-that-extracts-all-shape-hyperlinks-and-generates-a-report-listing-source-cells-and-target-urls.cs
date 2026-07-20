// Title: C# – Extract Shape Hyperlinks and Generate a Hyperlink Report Sheet with Aspose.Cells
// Description: Loads an Excel workbook, iterates through all worksheets and shapes, captures each shape’s upper‑left cell address and hyperlink URL, writes the sheet name, cell reference and URL to a new worksheet called ShapeHyperlinkReport, and saves the updated file.
// Keywords: Aspose.Cells | C# shape hyperlink extraction | Excel shape hyperlink report | extract shape URLs | hyperlink report worksheet | shape anchor cell | Aspose.Cells API | Excel automation | hyperlink audit | generate report sheet
// Common Searches: Aspose.Cells get hyperlink from shape | list all shape URLs in Excel using C# | create report of shape hyperlinks Aspose | extract shape hyperlink address C# | find cell of a shape Aspose.Cells | export shape hyperlinks to new sheet | C# Aspose.Cells shape hyperlink enumeration
// Developer Intent: Retrieve every hyperlink attached to shapes in a workbook and output a concise table of source cell locations and target URLs.
// Use Cases: Audit all clickable shapes across a workbook for compliance or documentation purposes. | Export shape hyperlink data before performing bulk updates or migrations. | Provide end‑users with a summary sheet that shows where each shape links, improving navigation.
// AI Prompts: Write a C# method using Aspose.Cells that scans all worksheets, finds shapes with hyperlinks, records the sheet name, upper‑left cell address, and URL into a new worksheet named 'ShapeHyperlinkReport', then saves the workbook. | Show how to extend the routine to also capture the shape name and hyperlink tooltip in the report. | Suggest a strategy for handling shapes that span multiple cells while still reporting the top‑left cell address and hyperlink.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace Example
{
    // Loads an Excel workbook, iterates through all worksheets and shapes, captures each shape’s upper‑left cell address and hyperlink URL, writes the sheet name, cell reference and URL to a new worksheet called ShapeHyperlinkReport, and saves the updated file.
    public class ShapeHyperlinkExtractor
    {
        // Extracts all shape hyperlinks from a workbook and creates a report worksheet.
        public static void ExtractShapeHyperlinks(string inputPath, string outputPath)
        {
            try
            {
                // Verify input file exists.
                if (!File.Exists(inputPath))
                    throw new FileNotFoundException($"Input file not found: {inputPath}");

                // Load the existing workbook.
                Workbook workbook = new Workbook(inputPath);

                // Add a new worksheet for the report.
                int reportIndex = workbook.Worksheets.Add();
                Worksheet reportSheet = workbook.Worksheets[reportIndex];
                reportSheet.Name = "ShapeHyperlinkReport";

                // Write header row.
                reportSheet.Cells[0, 0].PutValue("Source Cell");
                reportSheet.Cells[0, 1].PutValue("Target URL");

                int reportRow = 1; // Start after header.

                // Iterate through all worksheets.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all shapes in the worksheet.
                    foreach (Shape shape in sheet.Shapes)
                    {
                        // Check if the shape has a hyperlink.
                        Hyperlink link = shape.Hyperlink;
                        if (link != null && !string.IsNullOrEmpty(link.Address))
                        {
                            // Determine the cell where the shape is anchored (upper‑left corner).
                            int row = shape.UpperLeftRow;
                            int column = shape.UpperLeftColumn;

                            // Convert row/column to Excel cell name (e.g., A1).
                            string cellName = CellsHelper.CellIndexToName(row, column);

                            // Write the information to the report sheet.
                            reportSheet.Cells[reportRow, 0].PutValue($"{sheet.Name}!{cellName}");
                            reportSheet.Cells[reportRow, 1].PutValue(link.Address);
                            reportRow++;
                        }
                    }
                }

                // Save the workbook with the added report.
                workbook.Save(outputPath);
                Console.WriteLine($"Report saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the console application.
        public static void Main(string[] args)
        {
            string inputPath = "InputWorkbook.xlsx";
            string outputPath = "OutputWorkbook_WithReport.xlsx";

            // Allow overriding paths via command‑line arguments.
            if (args.Length >= 2)
            {
                inputPath = args[0];
                outputPath = args[1];
            }

            ExtractShapeHyperlinks(inputPath, outputPath);
        }
    }
}

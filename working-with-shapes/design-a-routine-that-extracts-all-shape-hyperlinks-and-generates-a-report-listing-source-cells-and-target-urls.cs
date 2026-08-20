// Title: Extract Shape Hyperlinks and Generate a Cell‑to‑URL Report with Aspose.Cells for .NET
// Description: Loads an Excel workbook, scans every worksheet for shapes with hyperlinks, determines each shape's top‑left cell (A1 style), and writes "Sheet!Cell -> URL" lines to a text report (optionally saving the workbook).
// Keywords: Aspose.Cells shape hyperlink extraction | C# list shape URLs in Excel | generate hyperlink report Aspose.Cells | shape anchor cell address .NET | export shape hyperlink data | Excel shape hyperlink audit
// Common Searches: Aspose.Cells get hyperlink from shape | C# extract all shape URLs in Excel workbook | report shape hyperlink addresses with cell reference | list shapes with hyperlinks Aspose.Cells .NET | export shape hyperlink mapping to text file
// Developer Intent: Create a text (or CSV) report that maps each shape’s anchored cell to its hyperlink URL across all worksheets.
// Use Cases: Compliance audit of all clickable shapes and their target URLs. | Automated documentation linking where shapes act as placeholders for external resources. | Pre‑release validation to ensure shape hyperlinks point only to approved domains.
// AI Prompts: Write C# code using Aspose.Cells that iterates through every shape in a workbook and outputs the shape's hyperlink address together with its anchored cell reference. | Provide a method that returns a Dictionary<string,string> where the key is "SheetName!CellAddress" of a shape and the value is the hyperlink URL, using Aspose.Cells for .NET. | Generate a sample script that saves the shape hyperlink report to a CSV file, including columns for worksheet, cell address, and URL.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeHyperlinkReport
{
    // Loads an Excel workbook, scans every worksheet for shapes with hyperlinks, determines each shape's top‑left cell (A1 style), and writes "Sheet!Cell -> URL" lines to a text report (optionally saving the workbook).
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("InputWorkbook.xlsx");

            // StringBuilder to collect the report lines
            StringBuilder reportBuilder = new StringBuilder();

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes on the current worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // Get the hyperlink associated with the shape (if any)
                    Hyperlink hyperlink = shape.Hyperlink;

                    // Proceed only when a hyperlink exists and has an address
                    if (hyperlink != null && !string.IsNullOrEmpty(hyperlink.Address))
                    {
                        // Determine the cell address where the shape is anchored
                        // UpperLeftRow and UpperLeftColumn give the top‑left cell indices (zero‑based)
                        int rowIndex = shape.UpperLeftRow;
                        int columnIndex = shape.UpperLeftColumn;
                        string cellAddress = CellsHelper.CellIndexToName(rowIndex, columnIndex);

                        // Build a line: SheetName!CellAddress -> HyperlinkAddress
                        string line = $"{sheet.Name}!{cellAddress} -> {hyperlink.Address}";
                        reportBuilder.AppendLine(line);
                    }
                }
            }

            // Write the report to a text file
            File.WriteAllText("ShapeHyperlinksReport.txt", reportBuilder.ToString());

            // Optionally, save the workbook (unchanged) to a new file
            workbook.Save("ProcessedWorkbook.xlsx");
        }
    }
}

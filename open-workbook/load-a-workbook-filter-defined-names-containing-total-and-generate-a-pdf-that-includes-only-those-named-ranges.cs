// Title: C# – Filter Named Ranges Containing “Total” and Export to PDF with Aspose.Cells
// Description: Load an Excel workbook, retrieve all defined names, select those whose text includes "Total" (case‑insensitive), copy each referenced range to a separate worksheet in a new workbook while preserving values and styles, and save the result as a PDF that contains only the filtered named ranges.
// Keywords: Aspose.Cells | C# | filter named ranges | export to PDF | Excel defined names | named range Total | Workbook to PDF | PdfSaveOptions | copy range to new sheet | Excel automation
// Common Searches: Aspose.Cells filter named ranges by keyword | C# export selected defined names to PDF | How to copy named ranges to new workbook Aspose.Cells | Generate PDF from specific Excel ranges C# | Save only Total named ranges as PDF using Aspose
// Developer Intent: Create a PDF that includes only the named ranges whose names contain the word "Total".
// Use Cases: Produce a financial summary PDF that shows only total rows defined as named ranges. | Generate a concise report PDF by extracting sections (totals, subtotals) identified through named ranges. | Automate archival of calculation results by exporting selected named ranges from a template workbook to a PDF.
// AI Prompts: Write C# code with Aspose.Cells to filter defined names by a keyword and export the matching ranges to a PDF. | Explain how to copy multiple named ranges into separate worksheets while preserving formatting before saving as PDF using Aspose.Cells. | Suggest improvements for handling overlapping named ranges so each appears on its own sheet in the generated PDF.

using System;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Alias to avoid conflict with System.Range
using AsposeRange = Aspose.Cells.Range;

// Load an Excel workbook, retrieve all defined names, select those whose text includes "Total" (case‑insensitive), copy each referenced range to a separate worksheet in a new workbook while preserving values and styles, and save the result as a PDF that contains only the filtered named ranges.
class FilterNamesToPdf
{
    static void Main()
    {
        try
        {
            // Path to the source workbook
            string sourcePath = "input.xlsx";

            // Verify that the source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook
            Workbook srcWorkbook = new Workbook(sourcePath);

            // Get all defined names in the workbook
            Name[] allNames = srcWorkbook.Worksheets.Names.Filter(NameScopeType.All, -1);

            // Select names that contain the word "Total" (case‑insensitive)
            var totalNames = allNames
                .Where(n => n.Text.IndexOf("Total", StringComparison.OrdinalIgnoreCase) >= 0)
                .ToArray();

            // Create a new workbook that will hold only the selected named ranges
            Workbook destWorkbook = new Workbook();
            // Remove the default sheet created by the constructor
            destWorkbook.Worksheets.Clear();

            // Copy each named range into its own worksheet in the destination workbook
            foreach (Name name in totalNames)
            {
                // Get all ranges referred by this name
                AsposeRange[] ranges = name.GetRanges();

                foreach (AsposeRange srcRange in ranges)
                {
                    // Add a new worksheet named after the defined name
                    Worksheet ws = destWorkbook.Worksheets.Add(name.Text);

                    // Copy cell values and styles from the source range to the new worksheet
                    int rowCount = srcRange.RowCount;
                    int colCount = srcRange.ColumnCount;

                    for (int i = 0; i < rowCount; i++)
                    {
                        for (int j = 0; j < colCount; j++)
                        {
                            Cell srcCell = srcRange[i, j];
                            Cell destCell = ws.Cells[i, j];

                            // Copy the cell value
                            destCell.PutValue(srcCell.Value);

                            // Copy the cell style
                            destCell.SetStyle(srcCell.GetStyle());
                        }
                    }

                    // Optionally assign the same name to the copied range in the new sheet
                    ws.Cells.CreateRange(0, 0, rowCount, colCount).Name = name.Text;
                }
            }

            // Save the resulting workbook as a PDF containing only the filtered named ranges
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            destWorkbook.Save("FilteredTotals.pdf", pdfOptions);

            Console.WriteLine("PDF generated successfully: FilteredTotals.pdf");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

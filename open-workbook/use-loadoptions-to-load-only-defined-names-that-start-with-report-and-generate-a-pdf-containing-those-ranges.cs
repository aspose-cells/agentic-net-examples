// Title: C# – Load Defined Names Starting with “Report” and Export Their Ranges to PDF with Aspose.Cells
// Description: Demonstrates how to use Aspose.Cells for .NET LoadOptions to load only cell data and defined names, filter names whose text begins with "Report", copy each referenced range into a new worksheet, and save the consolidated result as a PDF file.
// Keywords: Aspose.Cells | LoadOptions | defined names | named ranges | prefix filter | Report | C# .NET | export to PDF | copy range | Excel to PDF | filter named ranges
// Common Searches: Aspose.Cells load only defined names with prefix | C# filter named ranges starting with Report | Export selected named ranges to PDF using Aspose.Cells | How to copy named ranges to a new workbook in .NET | Generate PDF from specific defined names in Excel
// Developer Intent: Load only defined names that begin with "Report" and create a PDF containing the data from those ranges.
// Use Cases: Create a single printable PDF that summarizes all report sections (e.g., ReportSales, ReportInventory) defined in a financial workbook. | Automate extraction of specific report ranges for distribution to stakeholders as one consolidated PDF. | Generate archival PDFs of named report ranges to satisfy compliance or record‑keeping requirements. | Produce a dashboard PDF that includes only the named ranges needed for executive review. | Combine multiple report worksheets into one PDF without loading the entire workbook into memory.
// AI Prompts: Provide C# code using Aspose.Cells to open an Excel file with LoadOptions that load only cell data and defined names, filter names that start with "Report", copy each range to a new worksheet, and save as PDF. | Rewrite the example to use async/await for file I/O and reduce memory usage when processing large workbooks. | Show how to add a styled header containing the defined name above each copied range in the generated PDF. | Explain how to modify the code to include only visible cells from each named range. | Suggest logging strategies for processed names and graceful handling of missing or invalid ranges.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsDefinedNamesPdf
{
    // Demonstrates how to use Aspose.Cells for .NET LoadOptions to load only cell data and defined names, filter names whose text begins with "Report", copy each referenced range into a new worksheet, and save the consolidated result as a PDF file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the source Excel file
                string sourcePath = "SourceWorkbook.xlsx";

                // Verify that the source file exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {Path.GetFullPath(sourcePath)}");
                    return;
                }

                // Load the workbook with only cell data and defined names
                LoadOptions loadOptions = new LoadOptions
                {
                    LoadFilter = new LoadFilter(LoadDataFilterOptions.CellData | LoadDataFilterOptions.DefinedNames)
                };

                using (Workbook sourceWorkbook = new Workbook(sourcePath, loadOptions))
                {
                    // Create a new workbook that will contain only the matching ranges
                    using (Workbook resultWorkbook = new Workbook())
                    {
                        // Remove the default sheet and add a fresh one
                        resultWorkbook.Worksheets.Clear();
                        Worksheet resultSheet = resultWorkbook.Worksheets.Add("ReportRanges");

                        int destRow = 0; // Row index where the next range will be placed

                        // Iterate through all defined names in the source workbook
                        foreach (Name definedName in sourceWorkbook.Worksheets.Names)
                        {
                            // Check if the defined name starts with "Report" (case‑insensitive)
                            if (definedName.Text.StartsWith("Report", StringComparison.OrdinalIgnoreCase))
                            {
                                // Get the range that the defined name refers to
                                Aspose.Cells.Range srcRange = definedName.GetRange();

                                // Determine size of the source range
                                int rowCount = srcRange.RowCount;
                                int colCount = srcRange.ColumnCount;

                                // Create a destination range on the result sheet
                                Aspose.Cells.Range destRange = resultSheet.Cells.CreateRange(destRow, 0, rowCount, colCount);

                                // Copy the source range into the destination range
                                srcRange.Copy(destRange);

                                // Write the name of the range to the right of the copied data
                                resultSheet.Cells[destRow, colCount + 1].PutValue(definedName.Text);

                                // Move the destination row pointer below the copied block (add one empty row as separator)
                                destRow += rowCount + 2;
                            }
                        }

                        // Save the result workbook as PDF
                        PdfSaveOptions pdfOptions = new PdfSaveOptions();
                        string pdfPath = "ReportRanges.pdf";
                        resultWorkbook.Save(pdfPath, pdfOptions);

                        Console.WriteLine($"PDF generated successfully at: {Path.GetFullPath(pdfPath)}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

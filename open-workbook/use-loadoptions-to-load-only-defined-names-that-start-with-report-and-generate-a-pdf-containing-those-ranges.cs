// Title: C# – Load Defined Names Starting with “Report” and Export Their Ranges to PDF using Aspose.Cells
// Description: Demonstrates how to create LoadOptions with a LoadFilter that loads only defined names and cell data, open an Excel workbook, select named ranges whose names begin with "Report", copy each range into a new worksheet, and save the assembled workbook as a PDF. The approach reduces memory usage and speeds up processing by loading only the required data.
// Keywords: Aspose.Cells LoadOptions | LoadFilter DefinedNames | named ranges PDF export C# | extract Report named ranges | Excel to PDF selective export | .NET performance optimization | load only defined names Aspose
// Common Searches: Aspose.Cells load only defined names with prefix | Export named ranges to PDF C# Aspose.Cells | LoadFilter DefinedNames CellData example | Create PDF from specific Excel named ranges | How to filter defined names when opening a workbook
// Developer Intent: Load an Excel workbook using LoadOptions that includes only defined names beginning with "Report" and generate a PDF containing those selected ranges.
// Use Cases: Produce a compact PDF summary of all report sections identified by named ranges in a financial model. | Automate generation of lightweight PDFs that contain only data blocks following a naming convention, minimizing memory consumption. | Provide end‑users with a printable PDF of specific report areas without loading the full workbook.
// AI Prompts: Write C# code with Aspose.Cells to open a workbook using LoadOptions that loads only defined names and cell data, then export ranges whose names start with "Report" to a PDF. | Show how to filter defined names by a prefix, copy each corresponding range into a new workbook, and save the result as a PDF using Aspose.Cells for .NET. | Explain why using LoadFilter with LoadDataFilterOptions.DefinedNames | LoadDataFilterOptions.CellData improves performance when exporting selected named ranges to PDF.

using System;
using System.Data;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to create LoadOptions with a LoadFilter that loads only defined names and cell data, open an Excel workbook, select named ranges whose names begin with "Report", copy each range into a new worksheet, and save the assembled workbook as a PDF. The approach reduces memory usage and speeds up processing by loading only the required data.
class ReportRangesToPdf
{
    static void Main()
    {
        try
        {
            // Path to the source Excel file
            string sourcePath = "SourceWorkbook.xlsx";

            // Ensure the source file exists
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException($"Source file not found: {sourcePath}");

            // ---------- Load only Defined Names and Cell Data ----------
            // Create LoadOptions and set a LoadFilter that loads defined names and cell data.
            LoadOptions loadOptions = new LoadOptions
            {
                LoadFilter = new LoadFilter(LoadDataFilterOptions.DefinedNames | LoadDataFilterOptions.CellData)
            };

            // Load the workbook with the specified options
            Workbook sourceWorkbook = new Workbook(sourcePath, loadOptions);

            // ---------- Create a new workbook to hold the selected ranges ----------
            Workbook resultWorkbook = new Workbook();

            // Remove the default empty sheet if it exists
            if (resultWorkbook.Worksheets.Count > 0)
                resultWorkbook.Worksheets.RemoveAt(0);

            // Iterate through all defined names in the source workbook
            foreach (Name definedName in sourceWorkbook.Worksheets.Names)
            {
                // Process only names that start with "Report"
                if (!string.IsNullOrEmpty(definedName.Text) &&
                    definedName.Text.StartsWith("Report", StringComparison.OrdinalIgnoreCase))
                {
                    // Get the range that the defined name refers to
                    Aspose.Cells.Range srcRange = definedName.GetRange();

                    // Export the range to a DataTable (preserves values and types)
                    DataTable dt = srcRange.ExportDataTable();

                    // Add a new worksheet to the result workbook; name it after the defined name
                    Worksheet destSheet = resultWorkbook.Worksheets.Add(definedName.Text);

                    // Import the DataTable into the new worksheet starting at cell A1
                    ImportDataTableToWorksheet(destSheet, dt);
                }
            }

            // Ensure there is at least one worksheet before saving to PDF
            if (resultWorkbook.Worksheets.Count == 0)
                resultWorkbook.Worksheets.Add("Empty");

            // ---------- Save the result workbook as PDF ----------
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            string outputPdf = "ReportRanges.pdf";
            resultWorkbook.Save(outputPdf, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputPdf}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Helper method to import a DataTable into a worksheet cell-by-cell
    private static void ImportDataTableToWorksheet(Worksheet sheet, DataTable dt)
    {
        try
        {
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    sheet.Cells[i, j].PutValue(dt.Rows[i][j]);
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to import DataTable to worksheet '{sheet.Name}': {ex.Message}");
        }
    }
}

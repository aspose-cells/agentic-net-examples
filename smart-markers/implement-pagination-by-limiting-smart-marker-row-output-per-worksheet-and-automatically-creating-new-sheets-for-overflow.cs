// Title: How to paginate smart marker rows in Aspose.Cells for .NET by cloning a template sheet and limiting rows per worksheet
// AI Prompts: Generate C# code that uses Aspose.Cells WorkbookDesigner to split a DataTable into 500‑row chunks, clone a template worksheet for each chunk, and process smart markers on each new sheet. | Create a reusable C# method that accepts a DataTable and a maximum‑rows‑per‑sheet parameter, paginates the data, clones the template sheet, and applies smart markers using Aspose.Cells. | Write C# logic to rename each generated worksheet (e.g., "Page 1"), remove the original smart‑marker template after pagination, and save the workbook as a new Excel file.
// Common Searches: aspnet paginate smart marker output into multiple worksheets | c# Aspose.Cells limit rows per sheet when using smart markers | how to clone template worksheet for each data page with Aspose.Cells | split large DataTable into several Excel sheets using smart markers in .NET | auto generate new worksheets for overflow rows in Aspose.Cells smart markers
// Tags: smart-marker pagination Aspose.Cells | worksheet cloning for smart markers | smart marker row limit per sheet | automatic sheet creation Aspose.Cells | data pagination in Excel using Aspose.Cells

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The example loads a workbook containing a smart‑marker template, creates a DataTable of employee records, calculates the number of worksheets needed based on a 500‑row limit per sheet, clones the template for each page, copies the appropriate subset of rows into a new DataTable, processes the smart markers on each cloned sheet, removes the original template sheet, and saves the paginated workbook as 'PaginatedResult.xlsx'.
class SmartMarkerPagination
{
    // Maximum number of data rows to output per worksheet
    const int MaxRowsPerSheet = 500;

    static void Main()
    {
        try
        {
            // Verify that the template file exists
            const string templatePath = "TemplateWithSmartMarkers.xlsx";
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            // Load the workbook that contains the smart marker template sheet
            Workbook workbook = new Workbook(templatePath);

            // Assume the template sheet is the first worksheet
            int templateSheetIndex = 0;
            Worksheet templateSheet = workbook.Worksheets[templateSheetIndex];

            // -------------------------------------------------
            // Prepare the data source (replace with real data)
            // -------------------------------------------------
            DataTable sourceTable = GetSampleData(); // method defined below

            // -------------------------------------------------
            // Calculate pagination parameters
            // -------------------------------------------------
            int totalRows = sourceTable.Rows.Count;
            int pageCount = (totalRows + MaxRowsPerSheet - 1) / MaxRowsPerSheet;

            // -------------------------------------------------
            // Process each page
            // -------------------------------------------------
            for (int pageIndex = 0; pageIndex < pageCount; pageIndex++)
            {
                try
                {
                    // Clone the template sheet for the current page
                    int pageSheetIndex = workbook.Worksheets.AddCopy(templateSheetIndex);
                    Worksheet pageSheet = workbook.Worksheets[pageSheetIndex];
                    pageSheet.Name = $"Page {pageIndex + 1}";

                    // Create a DataTable that contains only the rows for this page
                    DataTable pageTable = sourceTable.Clone(); // copy structure only
                    int startRow = pageIndex * MaxRowsPerSheet;
                    int endRow = Math.Min(startRow + MaxRowsPerSheet, totalRows);

                    for (int r = startRow; r < endRow; r++)
                    {
                        pageTable.ImportRow(sourceTable.Rows[r]);
                    }

                    // -------------------------------------------------
                    // Run the Smart Marker processor on the current sheet
                    // -------------------------------------------------
                    WorkbookDesigner designer = new WorkbookDesigner(workbook);
                    designer.SetDataSource(pageTable);
                    // Process all sheets (the bool indicates whether to clear smart markers after processing)
                    designer.Process(false);
                }
                catch (Exception exPage)
                {
                    Console.WriteLine($"Error processing page {pageIndex + 1}: {exPage.Message}");
                }
            }

            // Optionally remove the original template sheet
            workbook.Worksheets.RemoveAt(templateSheetIndex);

            // Save the resulting workbook
            const string resultPath = "PaginatedResult.xlsx";
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully to {resultPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Sample method to generate dummy data; replace with actual data retrieval logic
    static DataTable GetSampleData()
    {
        DataTable dt = new DataTable("Employees");
        dt.Columns.Add("Name", typeof(string));
        dt.Columns.Add("Department", typeof(string));
        dt.Columns.Add("Salary", typeof(double));

        // Populate with 1200 rows to demonstrate pagination
        for (int i = 1; i <= 1200; i++)
        {
            dt.Rows.Add($"Employee {i}", $"Dept {(i % 5) + 1}", 30000 + (i * 10));
        }

        return dt;
    }
}

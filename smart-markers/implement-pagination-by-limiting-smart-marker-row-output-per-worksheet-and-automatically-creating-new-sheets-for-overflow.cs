// Title: Aspose.Cells for .NET – Paginate Smart Marker Rows Across Multiple Worksheets
// Description: Build a template with range‑based smart markers, divide a large collection into 100‑row chunks, copy the template sheet for each chunk, bind the chunk as the data source, process only the current worksheet with WorkbookDesigner, and save the workbook so every sheet respects the row limit.
// Keywords: Aspose.Cells | C# | smart markers | pagination | row limit per sheet | multiple worksheets | WorkbookDesigner | range smart markers | Excel export | large data set | copy worksheet | data chunking
// Common Searches: Aspose.Cells paginate smart marker rows | C# split smart marker output into multiple sheets | limit rows per worksheet using Aspose.Cells | automatic sheet creation when smart marker data exceeds limit | range smart markers pagination example .NET
// Developer Intent: Create an Excel file where smart‑marker expansion stops after a set number of rows per sheet and additional sheets are generated automatically for the remaining records.
// Use Cases: Employee directory export with a maximum of 100 rows per worksheet. | Invoice list that starts a new sheet after a predefined number of line items. | Large analytical report that automatically paginates data across several sheets using smart markers.
// AI Prompts: Generate C# code with Aspose.Cells that paginates smart‑marker output into worksheets, allowing the row limit to be configured. | Show how to name each generated worksheet based on its page number while preserving the header row. | Explain how to modify the loop to use a different chunk size or to add a summary sheet after pagination.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPaginationDemo
{
    // Sample data class
    // Build a template with range‑based smart markers, divide a large collection into 100‑row chunks, copy the template sheet for each chunk, bind the chunk as the data source, process only the current worksheet with WorkbookDesigner, and save the workbook so every sheet respects the row limit.
    public class Employee
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
        public string Department { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // ---------- 1. Create a template workbook with smart markers ----------
                Workbook templateWb = new Workbook();
                Worksheet templateWs = templateWb.Worksheets[0];

                // Header row
                templateWs.Cells["A1"].PutValue("Name");
                templateWs.Cells["B1"].PutValue("Age");
                templateWs.Cells["C1"].PutValue("Department");

                // Smart marker row (will be expanded by WorkbookDesigner)
                templateWs.Cells["A2"].PutValue("&=Employees.Name");
                templateWs.Cells["B2"].PutValue("&=Employees.Age");
                templateWs.Cells["C2"].PutValue("&=Employees.Department");

                // Define the range that contains the smart markers
                // The range must be named "_CellsSmartMarkers" when using range smart markers
                Aspose.Cells.Range smRange = templateWs.Cells.CreateRange("A2:C2");
                smRange.Name = "_CellsSmartMarkers";

                // ---------- 2. Prepare a large data source ----------
                const int totalRows = 250;          // total number of data rows
                const int maxRowsPerSheet = 100;    // rows allowed per worksheet

                List<Employee> allEmployees = new List<Employee>();
                for (int i = 1; i <= totalRows; i++)
                {
                    allEmployees.Add(new Employee
                    {
                        Name = $"Employee {i}",
                        Age = 20 + (i % 30),
                        Department = $"Dept {(i % 5) + 1}"
                    });
                }

                // ---------- 3. Paginate the data across worksheets ----------
                // The first sheet will be the original template sheet.
                // Subsequent sheets are copies of the template sheet.
                int processedCount = 0;          // how many rows have been processed
                int currentSheetIndex = 0;       // index of the sheet being processed

                // The workbook that will hold the final result
                Workbook resultWb = templateWb;   // start with the template workbook

                while (processedCount < allEmployees.Count)
                {
                    // Determine the size of the current chunk
                    int remaining = allEmployees.Count - processedCount;
                    int chunkSize = Math.Min(maxRowsPerSheet, remaining);

                    // Extract the chunk of data for the current sheet
                    List<Employee> chunk = allEmployees.GetRange(processedCount, chunkSize);

                    // If this is not the first sheet, add a fresh copy of the template sheet
                    if (currentSheetIndex > 0)
                    {
                        // AddCopy creates a new sheet based on the original template (index 0)
                        resultWb.Worksheets.AddCopy(0);
                    }

                    // Process the smart markers on the current sheet
                    WorkbookDesigner designer = new WorkbookDesigner
                    {
                        Workbook = resultWb,
                        // Using range smart markers; LineByLine is obsolete but kept for compatibility
                        LineByLine = false
                    };

                    // Set the data source for the current chunk
                    designer.SetDataSource("Employees", chunk);

                    // Process only the current sheet (sheetIndex, isPreserved)
                    designer.Process(currentSheetIndex, true);

                    // Move to the next chunk and next sheet
                    processedCount += chunkSize;
                    currentSheetIndex++;
                }

                // ---------- 4. Save the paginated workbook ----------
                resultWb.Save("PaginatedOutput.xlsx");
                Console.WriteLine("Workbook saved successfully as PaginatedOutput.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPaginationDemo
{
    // Sample data class used as a data source for smart markers
    public class Employee
    {
        public string? Name { get; set; }
        public int Age { get; set; }
        public string? Department { get; set; }
    }

    public class Program
    {
        // Maximum number of data rows (excluding header) that a single worksheet may contain
        private const int MaxRowsPerSheet = 100; // adjust as needed

        public static void Main()
        {
            try
            {
                // -------------------------------------------------
                // 1. Load the template workbook that contains smart markers
                // -------------------------------------------------
                const string templatePath = "TemplateWithSmartMarkers.xlsx";
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file not found: {templatePath}");
                    return;
                }

                Workbook workbook = new Workbook(templatePath);

                // -------------------------------------------------
                // 2. Prepare sample data source (list of employees)
                // -------------------------------------------------
                List<Employee> employees = new List<Employee>();
                for (int i = 1; i <= 350; i++) // generate more rows than MaxRowsPerSheet to trigger pagination
                {
                    employees.Add(new Employee
                    {
                        Name = $"Employee {i}",
                        Age = 20 + (i % 30),
                        Department = $"Dept {(i % 5) + 1}"
                    });
                }

                // -------------------------------------------------
                // 3. Configure WorkbookDesigner
                // -------------------------------------------------
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook,
                    // Use range smart markers (LineByLine = false) and define the range that contains the markers
                    LineByLine = false // obsolete but kept for compatibility with the original logic
                };

                // Assume the smart markers are placed in the range A2:C2 of the first sheet
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells.CreateRange("A2:C2").Name = "_CellsSmartMarkers";

                // Set the data source for the smart marker table named "Employees"
                designer.SetDataSource("Employees", employees);

                // -------------------------------------------------
                // 4. Process the smart markers – this expands the range into rows
                // -------------------------------------------------
                designer.Process();

                // -------------------------------------------------
                // 5. Paginate the resulting rows: limit rows per sheet and create new sheets for overflow
                // -------------------------------------------------
                PaginateWorksheetRows(workbook, MaxRowsPerSheet);

                // -------------------------------------------------
                // 6. Save the final workbook
                // -------------------------------------------------
                const string resultPath = "PaginatedResult.xlsx";
                workbook.Save(resultPath);
                Console.WriteLine($"Workbook saved to {resultPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        /// <summary>
        /// Splits data rows of each worksheet into multiple sheets when the row count exceeds the specified limit.
        /// The first row of each sheet is treated as a header and copied to every new sheet.
        /// </summary>
        /// <param name="workbook">Workbook to paginate.</param>
        /// <param name="maxRows">Maximum number of data rows (excluding header) per sheet.</param>
        private static void PaginateWorksheetRows(Workbook workbook, int maxRows)
        {
            try
            {
                // Snapshot of original sheet count because new sheets will be added during pagination
                int originalSheetCount = workbook.Worksheets.Count;

                for (int i = 0; i < originalSheetCount; i++)
                {
                    Worksheet sourceSheet = workbook.Worksheets[i];
                    int maxDataRow = sourceSheet.Cells.MaxDataRow; // zero‑based index of the last row containing data
                    if (maxDataRow <= 0) continue; // no data rows

                    // Assume the first row (index 0) is the header row
                    int headerRowIndex = 0;
                    int dataStartRow = headerRowIndex + 1; // first data row
                    int totalDataRows = maxDataRow - headerRowIndex; // number of data rows

                    if (totalDataRows <= maxRows) continue; // fits within the limit, no pagination needed

                    int rowsRemaining = totalDataRows;
                    int currentSourceRow = dataStartRow;

                    while (rowsRemaining > maxRows)
                    {
                        // Add a new worksheet and obtain its reference
                        int newSheetIndex = workbook.Worksheets.Add();
                        Worksheet newSheet = workbook.Worksheets[newSheetIndex];
                        newSheet.Name = $"{sourceSheet.Name}_Part{newSheetIndex + 1}";

                        // Copy the header row to the new sheet (row 0)
                        newSheet.Cells.CopyRows(sourceSheet.Cells, headerRowIndex, 0, 1);

                        // Determine how many rows to copy in this chunk
                        int rowsToCopy = Math.Min(maxRows, rowsRemaining);

                        // Copy the data rows
                        newSheet.Cells.CopyRows(sourceSheet.Cells, currentSourceRow, dataStartRow, rowsToCopy);

                        // Update counters for the next iteration
                        currentSourceRow += rowsToCopy;
                        rowsRemaining -= rowsToCopy;
                    }

                    // Delete excess rows from the original sheet, keeping only the header and the first 'maxRows' data rows
                    int rowsToDelete = totalDataRows - maxRows;
                    if (rowsToDelete > 0)
                    {
                        sourceSheet.Cells.DeleteRows(dataStartRow + maxRows, rowsToDelete);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Pagination error: {ex.Message}");
            }
        }
    }
}
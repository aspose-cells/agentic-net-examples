// Title: C# – Process only the first 100 smart‑marker rows in Aspose.Cells using a Range
// Description: Shows how to limit Aspose.Cells smart‑marker evaluation to the first 100 data rows by creating a named Range ("_CellsSmartMarkers") and calling WorkbookDesigner.Process(range, true). The sample builds a workbook with 200 marker rows, a data source of 150 employees, and saves the optimized result.
// Keywords: Aspose.Cells | smart markers | Range object | WorkbookDesigner.Process | C# | .NET | performance optimization | limit rows | named range | "_CellsSmartMarkers" | Excel template | data source | employee list
// Common Searches: Aspose.Cells limit smart marker rows | WorkbookDesigner process specific range C# | How to restrict smart markers to first 100 rows in Aspose.Cells | Using Range with smart markers Aspose.Cells | Performance tips for large smart marker worksheets
// Developer Intent: Restrict smart‑marker processing to a defined row range to boost performance.
// Use Cases: Generate a report that only needs the top 100 records while the template contains placeholders for many more rows. | Create a quick preview of a large dataset by processing smart markers in a limited range, reducing execution time. | Skip hidden or extra rows in a worksheet template by defining a named range that includes only the required smart markers.
// AI Prompts: Provide C# code that processes Aspose.Cells smart markers only within a specific Range and preserves unrecognized markers. | Explain how naming a Range "_CellsSmartMarkers" and passing it to WorkbookDesigner.Process(range, true) limits smart‑marker evaluation. | Show how to calculate zero‑based row and column indices when creating a Range for smart‑marker processing in Aspose.Cells.

using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

namespace SmartMarkerRangeDemo
{
    // Shows how to limit Aspose.Cells smart‑marker evaluation to the first 100 data rows by creating a named Range ("_CellsSmartMarkers") and calling WorkbookDesigner.Process(range, true). The sample builds a workbook with 200 marker rows, a data source of 150 employees, and saves the optimized result.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // ------------------------------------------------------------
                // Set up sample smart markers in the first 200 rows (A2:B201)
                // ------------------------------------------------------------
                // Header row
                sheet.Cells["A1"].PutValue("Name");
                sheet.Cells["B1"].PutValue("Value");

                // Populate smart markers for 200 rows
                for (int i = 2; i <= 201; i++)
                {
                    // Smart marker syntax: &=$DataSource.ColumnName
                    sheet.Cells[i - 1, 0].PutValue("&=$Employees.Name");
                    sheet.Cells[i - 1, 1].PutValue("&=$Employees.Salary");
                }

                // ------------------------------------------------------------
                // Prepare a data source with more than 100 records
                // ------------------------------------------------------------
                ArrayList employees = new ArrayList();
                for (int i = 1; i <= 150; i++)
                {
                    employees.Add(new Employee
                    {
                        Name = $"Employee {i}",
                        Salary = 3000 + i * 10
                    });
                }

                // ------------------------------------------------------------
                // Configure WorkbookDesigner
                // ------------------------------------------------------------
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };
                designer.SetDataSource("Employees", employees);

                // ------------------------------------------------------------
                // Define a range that covers only the first 100 data rows
                // Rows are zero‑based, so row index 1 corresponds to Excel row 2.
                // CreateRange(startRow, startColumn, totalRows, totalColumns)
                // ------------------------------------------------------------
                int startRow = 1;          // Excel row 2 (first data row)
                int startColumn = 0;       // Column A
                int totalRows = 100;       // Process only first 100 rows
                int totalColumns = 2;      // Columns A and B

                // Resolve ambiguity with System.Range by using fully qualified name
                Aspose.Cells.Range smartMarkerRange = sheet.Cells.CreateRange(startRow, startColumn, totalRows, totalColumns);
                // The name "_CellsSmartMarkers" tells the designer that this range contains smart markers
                smartMarkerRange.Name = "_CellsSmartMarkers";

                // ------------------------------------------------------------
                // Process only the defined range (true = preserve unrecognized markers)
                // ------------------------------------------------------------
                designer.Process(smartMarkerRange, true);

                // ------------------------------------------------------------
                // Save the result
                // ------------------------------------------------------------
                string outputPath = "SmartMarkersFirst100Rows.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Simple POCO class used as a data source
        public class Employee
        {
            public string Name { get; set; } = string.Empty;
            public double Salary { get; set; }
        }
    }
}

// Title: Limit Smart Marker Evaluation to 100 Rows Using Range in Aspose.Cells for .NET (C#)
// Description: A C# sample that builds a workbook, inserts smart markers, binds a DataTable, creates a Range covering rows 0‑99 and all populated columns, and calls WorkbookDesigner.Process with that range to evaluate only that segment, boosting performance.
// Keywords: Aspose.Cells C# range smart markers | process smart markers specific rows | WorkbookDesigner limited range | smart marker performance .NET | restrict smart marker evaluation | Aspose.Cells Range object | C# Excel smart markers | limit rows Aspose.Cells
// Common Searches: Aspose.Cells process smart markers only first rows | C# WorkbookDesigner Process with range parameter | How to limit smart marker evaluation in Aspose.Cells | Range object for smart markers performance .NET | Smart markers processing subset of cells Aspose
// Developer Intent: Execute smart‑marker replacement solely within a defined cell block to reduce execution time.
// Use Cases: Generating a summary sheet where only the top section contains placeholders. | Working with large templates that have many markers but only a particular data block needs population. | Accelerating report creation by processing a bounded area while leaving other markers untouched.
// AI Prompts: Provide a C# snippet that processes smart markers in a specified Range using Aspose.Cells. | Show how to call WorkbookDesigner.Process(range, true) to evaluate markers in rows 0‑99. | Explain how to preserve unknown smart markers while limiting processing to a cell region.

using System;
using System.Data;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace SmartMarkerRangeExample
{
    // A C# sample that builds a workbook, inserts smart markers, binds a DataTable, creates a Range covering rows 0‑99 and all populated columns, and calls WorkbookDesigner.Process with that range to evaluate only that segment, boosting performance.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook (or load a template containing smart markers)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // ------------------------------------------------------------
                // Sample smart markers placed in the first 100 rows for demo.
                // In a real scenario the template would already contain them.
                // ------------------------------------------------------------
                for (int i = 0; i < 100; i++)
                {
                    // Example smart marker that will be populated from a data source named "Data"
                    sheet.Cells[i, 0].PutValue($"&=Data.Column{i + 1}");
                }

                // Set up a data source with at least 100 columns (dummy data for illustration)
                DataTable dataTable = new DataTable("Data");
                for (int i = 0; i < 100; i++)
                {
                    dataTable.Columns.Add($"Column{i + 1}", typeof(string));
                }

                DataRow row = dataTable.NewRow();
                for (int i = 0; i < 100; i++)
                {
                    row[i] = $"Value {i + 1}";
                }
                dataTable.Rows.Add(row);

                // Initialize WorkbookDesigner and assign the workbook
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };

                // Register the data source
                designer.SetDataSource(dataTable);

                // ------------------------------------------------------------
                // Restrict processing to the first 100 rows using a Range object.
                // The range starts at row 0, column 0 and spans 100 rows.
                // ------------------------------------------------------------
                int startRow = 0;
                int startColumn = 0;
                int totalRows = 100; // first 100 rows
                int totalColumns = sheet.Cells.MaxDataColumn + 1; // include all columns that have data

                AsposeRange smartMarkerRange = sheet.Cells.CreateRange(startRow, startColumn, totalRows, totalColumns);
                // Naming the range as required for smart marker processing (optional but common)
                smartMarkerRange.Name = "_CellsSmartMarkers";

                // Process only the defined range; true = preserve unrecognized markers
                designer.Process(smartMarkerRange, true);

                // Save the resulting workbook
                string outputPath = "SmartMarkersProcessedFirst100Rows.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

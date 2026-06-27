using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    class WorkbookConnectionsCsvReport
    {
        static void Main()
        {
            // Folder containing the Excel files to process
            string folderPath = @"C:\ExcelFiles";

            // Output CSV file path
            string outputCsvPath = Path.Combine(folderPath, "WorkbookConnectionsReport.csv");

            // Create a new workbook that will hold the report data
            Workbook reportWorkbook = new Workbook();

            // Get the first (and only) worksheet in the report workbook
            Worksheet sheet = reportWorkbook.Worksheets[0];

            // Write header row
            sheet.Cells["A1"].PutValue("Workbook Name");
            sheet.Cells["B1"].PutValue("Connection Type");
            sheet.Cells["C1"].PutValue("Connection Name");

            int currentRow = 1; // zero‑based index; row 1 is the second row (after header)

            // Iterate through all files in the specified folder
            foreach (string filePath in Directory.GetFiles(folderPath))
            {
                // Consider only Excel files (you can extend the filter as needed)
                string extension = Path.GetExtension(filePath).ToLowerInvariant();
                if (extension != ".xlsx" && extension != ".xls" && extension != ".xlsm" && extension != ".xlsb")
                    continue;

                // Load the workbook using the provided constructor (load rule)
                Workbook wb = new Workbook(filePath);

                // Access the collection of external data connections (property rule)
                ExternalConnectionCollection connections = wb.DataConnections;

                // If there are no connections, still record the workbook name with empty fields
                if (connections.Count == 0)
                {
                    sheet.Cells[currentRow, 0].PutValue(Path.GetFileName(filePath));
                    sheet.Cells[currentRow, 1].PutValue(string.Empty);
                    sheet.Cells[currentRow, 2].PutValue(string.Empty);
                    currentRow++;
                }
                else
                {
                    // Iterate through each connection and write its details
                    for (int i = 0; i < connections.Count; i++)
                    {
                        ExternalConnection conn = connections[i];

                        // Workbook name (file name only)
                        sheet.Cells[currentRow, 0].PutValue(Path.GetFileName(filePath));

                        // Connection type – use the runtime type name
                        sheet.Cells[currentRow, 1].PutValue(conn.GetType().Name);

                        // Connection name – most connection types expose a Name property
                        sheet.Cells[currentRow, 2].PutValue(conn.Name);

                        currentRow++;
                    }
                }

                // Dispose the loaded workbook (optional, as it implements IDisposable)
                wb.Dispose();
            }

            // Save the report workbook as CSV using the provided Save method (save rule)
            reportWorkbook.Save(outputCsvPath, SaveFormat.Csv);

            // Clean up
            reportWorkbook.Dispose();

            Console.WriteLine($"Connection report generated at: {outputCsvPath}");
        }
    }
}
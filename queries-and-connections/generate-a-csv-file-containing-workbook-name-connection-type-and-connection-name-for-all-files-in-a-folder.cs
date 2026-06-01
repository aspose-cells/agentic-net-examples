using System;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class WorkbookConnectionsCsvGenerator
{
    static void Main()
    {
        // Folder containing the Excel files
        string folderPath = @"C:\InputWorkbooks";

        // Output CSV file path
        string outputCsvPath = @"C:\Output\WorkbookConnections.csv";

        // Create a new workbook that will hold the CSV data
        Workbook csvWorkbook = new Workbook(); // create rule
        Worksheet sheet = csvWorkbook.Worksheets[0];

        // Write CSV header
        int currentRow = 0;
        sheet.Cells[currentRow, 0].PutValue("Workbook Name");
        sheet.Cells[currentRow, 1].PutValue("Connection Type");
        sheet.Cells[currentRow, 2].PutValue("Connection Name");
        currentRow++;

        // Get all files in the folder (you can filter extensions as needed)
        string[] files = Directory.GetFiles(folderPath)
                                  .Where(f => f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                                              f.EndsWith(".xls", StringComparison.OrdinalIgnoreCase) ||
                                              f.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase) ||
                                              f.EndsWith(".xlsb", StringComparison.OrdinalIgnoreCase))
                                  .ToArray();

        foreach (string filePath in files)
        {
            // Load each workbook (load rule)
            Workbook wb = new Workbook(filePath);

            // Iterate through its external data connections
            foreach (ExternalConnection conn in wb.DataConnections)
            {
                // Connection type (class name) and connection name
                string connectionType = conn.GetType().Name;
                string connectionName = conn.Name;

                // Write a row to the CSV workbook
                sheet.Cells[currentRow, 0].PutValue(Path.GetFileName(filePath));
                sheet.Cells[currentRow, 1].PutValue(connectionType);
                sheet.Cells[currentRow, 2].PutValue(connectionName);
                currentRow++;
            }
        }

        // Save the result as CSV (save rule)
        csvWorkbook.Save(outputCsvPath, SaveFormat.Csv);
    }
}
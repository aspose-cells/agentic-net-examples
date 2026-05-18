using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class Program
{
    static void Main()
    {
        // Folder containing the Excel files
        string folderPath = @"C:\ExcelFiles";

        // Path for the output CSV file
        string csvPath = Path.Combine(folderPath, "WorkbookConnections.csv");

        // Create CSV and write header
        using (var writer = new StreamWriter(csvPath))
        {
            writer.WriteLine("WorkbookName,ConnectionType,ConnectionName");

            // Iterate through all .xlsx files in the folder
            foreach (string filePath in Directory.GetFiles(folderPath, "*.xlsx"))
            {
                // Load the workbook (uses the Workbook(string) constructor rule)
                Workbook workbook = new Workbook(filePath);

                string workbookName = Path.GetFileName(filePath);

                // Access the external connections collection
                ExternalConnectionCollection connections = workbook.DataConnections;

                // Write each connection's details to the CSV
                for (int i = 0; i < connections.Count; i++)
                {
                    ExternalConnection conn = connections[i];
                    string connectionType = conn.GetType().Name;   // e.g., WebQueryConnection, OdbcConnection, etc.
                    string connectionName = conn.Name;            // Name of the connection

                    writer.WriteLine($"{workbookName},{connectionType},{connectionName}");
                }
            }
        }

        Console.WriteLine($"CSV file generated at: {csvPath}");
    }
}
using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExternalConnectionDemo
{
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // Scenario 1: Load an existing XLSX workbook that already contains
            // external data connections and inspect / modify its connection settings.
            // -----------------------------------------------------------------
            string sourcePath = "SourceWithConnections.xlsx";

            // Ensure the source file exists; create an empty workbook if it does not.
            if (!File.Exists(sourcePath))
            {
                var emptyWb = new Workbook();
                emptyWb.Save(sourcePath, SaveFormat.Xlsx);
            }

            Workbook wb = new Workbook(sourcePath);

            // Access the collection of external connections
            ExternalConnectionCollection connections = wb.DataConnections;

            Console.WriteLine($"Number of external connections: {connections.Count}");

            for (int i = 0; i < connections.Count; i++)
            {
                ExternalConnection conn = connections[i];

                // Display key properties of each connection
                Console.WriteLine($"--- Connection {i + 1} ---");
                Console.WriteLine($"Name                     : {conn.Name}");
                Console.WriteLine($"Class Type               : {conn.ClassType}");
                Console.WriteLine($"Source Type              : {conn.SourceType}");
                Console.WriteLine($"OdcFile (external file)  : {conn.OdcFile}");
                Console.WriteLine($"OnlyUseConnectionFile    : {conn.OnlyUseConnectionFile}");
                Console.WriteLine($"BackgroundRefresh        : {conn.BackgroundRefresh}");
                Console.WriteLine($"RefreshOnLoad            : {conn.RefreshOnLoad}");
                Console.WriteLine();

                // Example modification: force the connection to use only the ODC file
                conn.OnlyUseConnectionFile = true;

                // Example modification: point to a different ODC file
                conn.OdcFile = @"C:\ExternalConnections\NewConnection.odc";
            }

            // -----------------------------------------------------------------
            // Scenario 2: Add an external link to another workbook.
            // This demonstrates linking to data in a separate XLSX file.
            // -----------------------------------------------------------------
            string externalFile = "ExternalData.xlsx";
            string[] linkedSheets = new string[] { "Sheet1!A1:B10", "Sheet2!C5" };

            // Add the external link to the workbook's worksheets collection
            int linkIndex = wb.Worksheets.ExternalLinks.Add(externalFile, linkedSheets);
            Console.WriteLine($"External link added at index {linkIndex}.");

            // Optionally, modify the data source of the added link
            wb.Worksheets.ExternalLinks[linkIndex].DataSource = @"D:\DataSources\UpdatedExternalData.xlsx";

            // -----------------------------------------------------------------
            // Scenario 3: Save the modified workbook using the default XLSX format.
            // -----------------------------------------------------------------
            string outputPath = "ModifiedWorkbook.xlsx";
            wb.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}
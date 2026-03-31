using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    public class ODataConnectionDemo
    {
        public static void Run()
        {
            // Load an existing workbook that contains an OData (Data Feed) connection.
            // The workbook path should be replaced with an actual file path.
            string inputPath = "ODataSample.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all external data connections in the workbook.
            foreach (ExternalConnection conn in workbook.DataConnections)
            {
                // Identify OData connections by checking the source type.
                // OData connections are represented by the DataFeedDataModel enum value.
                if (conn.SourceType == ConnectionDataSourceType.DataFeedDataModel)
                {
                    Console.WriteLine("=== OData Connection Found ===");
                    Console.WriteLine($"Name                : {conn.Name}");
                    Console.WriteLine($"Source Type         : {conn.SourceType}");
                    Console.WriteLine($"Command (URL)       : {conn.Command}");
                    Console.WriteLine($"OdcFile (if any)    : {conn.OdcFile}");
                    Console.WriteLine($"Refresh On Load     : {conn.RefreshOnLoad}");
                    Console.WriteLine($"Background Refresh  : {conn.BackgroundRefresh}");
                    Console.WriteLine();
                }
            }

            // Save the workbook in the default XLSX format.
            // The output path can be changed as needed.
            string outputPath = "ODataConnectionResult.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }

        public static void Main()
        {
            Run();
        }
    }
}
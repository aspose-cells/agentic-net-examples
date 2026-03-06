using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExternalConnectionDemo
{
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Load an XLSX workbook, read its external connections,
            //    modify a property and save it back as XLSX.
            // -----------------------------------------------------------------
            string xlsxInputPath = "input.xlsx";
            string xlsxOutputPath = "output.xlsx";

            // Load the XLSX workbook (uses Workbook(string) constructor)
            Workbook xlsxWorkbook = new Workbook(xlsxInputPath);

            // Access the external connections collection
            ExternalConnectionCollection xlsxConnections = xlsxWorkbook.DataConnections;

            // Iterate through each connection and display the OnlyUseConnectionFile property
            for (int i = 0; i < xlsxConnections.Count; i++)
            {
                ExternalConnection conn = xlsxConnections[i];
                Console.WriteLine($"XLSX Connection {i + 1}: OnlyUseConnectionFile = {conn.OnlyUseConnectionFile}");

                // Example modification: toggle the property
                conn.OnlyUseConnectionFile = !conn.OnlyUseConnectionFile;
                Console.WriteLine($"XLSX Connection {i + 1}: OnlyUseConnectionFile toggled to {conn.OnlyUseConnectionFile}");
            }

            // Save the modified workbook as XLSX (uses Save(string, SaveFormat))
            xlsxWorkbook.Save(xlsxOutputPath, SaveFormat.Xlsx);
            Console.WriteLine($"XLSX workbook saved to '{xlsxOutputPath}'.");
            Console.WriteLine();

            // -----------------------------------------------------------------
            // 2. Load an XLS workbook, read/write its external connections,
            //    and save it using XlsSaveOptions.
            // -----------------------------------------------------------------
            string xlsInputPath = "sample.xls";
            string xlsOutputPath = "sample_modified.xls";

            // Load the XLS workbook
            Workbook xlsWorkbook = new Workbook(xlsInputPath);

            // Access external connections
            ExternalConnectionCollection xlsConnections = xlsWorkbook.DataConnections;

            // Example: set OnlyUseConnectionFile to true for all connections
            for (int i = 0; i < xlsConnections.Count; i++)
            {
                ExternalConnection conn = xlsConnections[i];
                Console.WriteLine($"XLS Connection {i + 1}: Original OnlyUseConnectionFile = {conn.OnlyUseConnectionFile}");
                conn.OnlyUseConnectionFile = true;
                Console.WriteLine($"XLS Connection {i + 1}: Updated OnlyUseConnectionFile = {conn.OnlyUseConnectionFile}");
            }

            // Create save options for XLS format
            XlsSaveOptions xlsSaveOptions = new XlsSaveOptions();

            // Save the workbook with the specified options (uses Save(string, SaveOptions))
            xlsWorkbook.Save(xlsOutputPath, xlsSaveOptions);
            Console.WriteLine($"XLS workbook saved to '{xlsOutputPath}'.");
            Console.WriteLine();

            // -----------------------------------------------------------------
            // 3. Load an XLSB workbook, read/write its external connections,
            //    and save it using XlsbSaveOptions.
            // -----------------------------------------------------------------
            string xlsbInputPath = "sample.xlsb";
            string xlsbOutputPath = "sample_modified.xlsb";

            // Load the XLSB workbook
            Workbook xlsbWorkbook = new Workbook(xlsbInputPath);

            // Access external connections
            ExternalConnectionCollection xlsbConnections = xlsbWorkbook.DataConnections;

            // Example: set OnlyUseConnectionFile to false for all connections
            for (int i = 0; i < xlsbConnections.Count; i++)
            {
                ExternalConnection conn = xlsbConnections[i];
                Console.WriteLine($"XLSB Connection {i + 1}: Original OnlyUseConnectionFile = {conn.OnlyUseConnectionFile}");
                conn.OnlyUseConnectionFile = false;
                Console.WriteLine($"XLSB Connection {i + 1}: Updated OnlyUseConnectionFile = {conn.OnlyUseConnectionFile}");
            }

            // Create save options for XLSB format
            XlsbSaveOptions xlsbSaveOptions = new XlsbSaveOptions();

            // Save the workbook with the specified options (uses Save(string, SaveOptions))
            xlsbWorkbook.Save(xlsbOutputPath, xlsbSaveOptions);
            Console.WriteLine($"XLSB workbook saved to '{xlsbOutputPath}'.");
        }
    }
}
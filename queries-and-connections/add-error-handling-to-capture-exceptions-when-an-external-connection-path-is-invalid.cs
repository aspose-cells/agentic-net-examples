using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    public class ExternalConnectionInvalidPathHandling
    {
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Attempt to work with external connections
            try
            {
                // Access the external connections collection
                ExternalConnectionCollection connections = workbook.DataConnections;

                // If there is at least one connection, try to set an invalid ODC file path
                if (connections.Count > 0)
                {
                    ExternalConnection conn = connections[0];

                    // Intentionally set an invalid path to trigger an exception
                    conn.OdcFile = @"Z:\NonExistentFolder\InvalidConnection.odc";
                }
                else
                {
                    Console.WriteLine("No external connections found in the workbook.");
                }
            }
            catch (Exception ex)
            {
                // Capture and display any exception related to an invalid external connection path
                Console.WriteLine($"Error while setting OdcFile: {ex.Message}");
            }

            // Attempt to add an external link with an invalid file path
            try
            {
                // Get the external links collection
                ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

                // Define an invalid file name (non‑existent file)
                string invalidFileName = @"C:\InvalidFolder\MissingWorkbook.xlsx";
                string[] sheetNames = new string[] { "Sheet1" };

                // Verify the file exists before attempting to add the link
                if (File.Exists(invalidFileName))
                {
                    int index = externalLinks.Add(invalidFileName, sheetNames);
                    Console.WriteLine($"External link added at index {index} (unexpected).");
                }
                else
                {
                    Console.WriteLine($"File not found: {invalidFileName}. Skipping external link addition.");
                }
            }
            catch (Exception ex)
            {
                // Capture and display any exception related to adding an invalid external link
                Console.WriteLine($"Error while adding external link: {ex.Message}");
            }

            // Save the workbook (lifecycle rule: save)
            try
            {
                workbook.Save("ExternalConnectionInvalidPathDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error while saving workbook: {ex.Message}");
            }
        }
    }
}
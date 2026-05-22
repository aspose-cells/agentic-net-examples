using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    public class EnumerateExternalQueryTables
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Path to the input workbook
            string inputPath = "InputWorkbookWithQueryTables.xlsx";

            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook that contains query tables
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Iterate through each worksheet in the workbook
            for (int wsIndex = 0; wsIndex < workbook.Worksheets.Count; wsIndex++)
            {
                Worksheet sheet = workbook.Worksheets[wsIndex];
                Console.WriteLine($"Worksheet: {sheet.Name}");

                // Check if the worksheet has any query tables
                if (sheet.QueryTables.Count == 0)
                {
                    Console.WriteLine("  No query tables in this worksheet.");
                    continue;
                }

                // Enumerate each query table
                for (int qtIndex = 0; qtIndex < sheet.QueryTables.Count; qtIndex++)
                {
                    QueryTable queryTable = sheet.QueryTables[qtIndex];
                    Console.WriteLine($"  Query Table {qtIndex + 1}: {queryTable.Name}");
                    Console.WriteLine($"    Result Range: {queryTable.ResultRange.Address}");

                    // Get the external connection associated with the query table
                    ExternalConnection extConn = queryTable.ExternalConnection;
                    if (extConn != null)
                    {
                        Console.WriteLine("    Linked to external data source:");
                        Console.WriteLine($"      Connection ID    : {extConn.Id}");
                        Console.WriteLine($"      Connection Name  : {extConn.Name}");
                        Console.WriteLine($"      Connection Type  : {extConn.ClassType}");
                        Console.WriteLine($"      Connection String: {extConn.ConnectionString}");
                        Console.WriteLine($"      Refresh on Load  : {extConn.RefreshOnLoad}");
                    }
                    else
                    {
                        Console.WriteLine("    No external connection associated with this query table.");
                    }
                }
            }

            // List workbook‑level external links (if any)
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;
            if (externalLinks.Count > 0)
            {
                Console.WriteLine("\nWorkbook External Links:");
                for (int i = 0; i < externalLinks.Count; i++)
                {
                    ExternalLink link = externalLinks[i];
                    Console.WriteLine($"  Link {i + 1}: DataSource = {link.DataSource}");
                }
            }
            else
            {
                Console.WriteLine("\nNo workbook external links found.");
            }

            // Save the workbook (unchanged) if needed
            try
            {
                workbook.Save("EnumeratedQueryTablesOutput.xlsx");
                Console.WriteLine("\nWorkbook saved as EnumeratedQueryTablesOutput.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
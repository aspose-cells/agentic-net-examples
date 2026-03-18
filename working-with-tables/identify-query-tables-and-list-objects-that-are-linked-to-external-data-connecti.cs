using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;
using Aspose.Cells.Tables;   // for ListObject

namespace AsposeCellsExamples
{
    public class IdentifyExternalDataLinks
    {
        public static void Run()
        {
            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook("input.xlsx");

            Console.WriteLine("=== Query Tables with External Connections ===");
            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Check for query tables
                if (sheet.QueryTables.Count > 0)
                {
                    Console.WriteLine($"Worksheet: {sheet.Name}");
                    for (int i = 0; i < sheet.QueryTables.Count; i++)
                    {
                        QueryTable qt = sheet.QueryTables[i];
                        Console.WriteLine($"  QueryTable Name: {qt.Name}");

                        // Get the related external connection (if any)
                        ExternalConnection conn = qt.ExternalConnection;
                        if (conn != null)
                        {
                            Console.WriteLine($"    External Connection Id: {conn.Id}");
                            Console.WriteLine($"    External Connection Name: {conn.Name}");
                            Console.WriteLine($"    External Connection Type: {conn.ClassType}");
                        }
                        else
                        {
                            Console.WriteLine("    No external connection associated with this query table.");
                        }
                    }
                }

                // Check for list objects (tables) that may have external connections
                if (sheet.ListObjects.Count > 0)
                {
                    Console.WriteLine($"Worksheet: {sheet.Name} (ListObjects)");
                    for (int i = 0; i < sheet.ListObjects.Count; i++)
                    {
                        ListObject lo = sheet.ListObjects[i];
                        // Use DisplayName as the table name (Name property not available in this version)
                        Console.WriteLine($"  ListObject Name: {lo.DisplayName}");

                        // Some ListObjects expose an ExternalConnection property
                        // Use reflection to safely access it if present
                        var prop = lo.GetType().GetProperty("ExternalConnection");
                        if (prop != null)
                        {
                            var loConn = prop.GetValue(lo) as ExternalConnection;
                            if (loConn != null)
                            {
                                Console.WriteLine($"    External Connection Id: {loConn.Id}");
                                Console.WriteLine($"    External Connection Name: {loConn.Name}");
                                Console.WriteLine($"    External Connection Type: {loConn.ClassType}");
                            }
                            else
                            {
                                Console.WriteLine("    No external connection associated with this list object.");
                            }
                        }
                        else
                        {
                            Console.WriteLine("    ListObject does not expose an ExternalConnection property.");
                        }
                    }
                }
            }

            Console.WriteLine("\n=== Workbook External Links ===");
            // External links stored in the workbook
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;
            for (int i = 0; i < externalLinks.Count; i++)
            {
                ExternalLink link = externalLinks[i];
                Console.WriteLine($"  External Link {i + 1}: DataSource = {link.DataSource}");
            }

            Console.WriteLine("\n=== Workbook Data Connections ===");
            // Data connections defined in the workbook
            ExternalConnectionCollection dataConns = workbook.DataConnections;
            for (int i = 0; i < dataConns.Count; i++)
            {
                ExternalConnection conn = dataConns[i];
                Console.WriteLine($"  Connection {i + 1}: Id = {conn.Id}, Name = {conn.Name}, Type = {conn.ClassType}");
            }

            // Save the workbook (lifecycle rule: save) – optional if you made changes
            workbook.Save("output.xlsx");
        }

        // Entry point
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
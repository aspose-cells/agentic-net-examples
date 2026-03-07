using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    public class FindExternalDataConnectionsDemo
    {
        public static void Run()
        {
            string inputPath = "InputWithExternalData.xlsx";

            // If the input file does not exist, create a blank workbook.
            if (!File.Exists(inputPath))
            {
                var wb = new Workbook();
                wb.Worksheets[0].Name = "Sheet1";
                wb.Save(inputPath);
            }

            // Load the workbook that may contain external data connections
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Console.WriteLine($"Worksheet: {sheet.Name}");

                // ----- QueryTables -----
                if (sheet.QueryTables.Count > 0)
                {
                    for (int i = 0; i < sheet.QueryTables.Count; i++)
                    {
                        QueryTable qt = sheet.QueryTables[i];
                        ExternalConnection conn = qt.ExternalConnection;

                        if (conn != null)
                        {
                            Console.WriteLine($"  QueryTable \"{qt.Name}\" is linked to external connection:");
                            Console.WriteLine($"    Connection Id   : {conn.Id}");
                            Console.WriteLine($"    Connection Name : {conn.Name}");
                            Console.WriteLine($"    Class Type      : {conn.ClassType}");
                            Console.WriteLine($"    Connection String: {conn.ConnectionString}");
                        }
                        else
                        {
                            Console.WriteLine($"  QueryTable \"{qt.Name}\" has no external connection.");
                        }
                    }
                }

                // ----- ListObjects (tables) -----
                if (sheet.ListObjects.Count > 0)
                {
                    for (int i = 0; i < sheet.ListObjects.Count; i++)
                    {
                        ListObject lo = sheet.ListObjects[i];
                        ExternalConnection conn = lo.QueryTable?.ExternalConnection;

                        if (conn != null)
                        {
                            Console.WriteLine($"  ListObject \"{lo.DisplayName}\" is linked to external connection:");
                            Console.WriteLine($"    Connection Id   : {conn.Id}");
                            Console.WriteLine($"    Connection Name : {conn.Name}");
                            Console.WriteLine($"    Class Type      : {conn.ClassType}");
                            Console.WriteLine($"    Connection String: {conn.ConnectionString}");
                        }
                        else
                        {
                            Console.WriteLine($"  ListObject \"{lo.DisplayName}\" has no external connection.");
                        }
                    }
                }
            }

            // List all external connections defined at the workbook level
            var allConns = workbook.DataConnections;
            Console.WriteLine($"Workbook contains {allConns.Count} external connection(s).");
            for (int i = 0; i < allConns.Count; i++)
            {
                var conn = allConns[i];
                Console.WriteLine($"  Connection {i + 1}: Id={conn.Id}, Name={conn.Name}, Type={conn.Type}, ConnectionString={conn.ConnectionString}");
            }

            // Save the workbook (no changes made, just to demonstrate the save rule)
            workbook.Save("OutputWithExternalDataInfo.xlsx");
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            FindExternalDataConnectionsDemo.Run();
        }
    }
}
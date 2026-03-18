using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    public class IdentifyLinkedObjects
    {
        public static void Run()
        {
            // Load the workbook (replace with your actual file path)
            Workbook workbook = new Workbook("InputWorkbook.xlsx");

            // Specify the name of the external data connection to search for
            string targetConnectionName = "MyExternalConnection";

            // Retrieve the external connection from the workbook's DataConnections collection
            ExternalConnection targetConnection = null;
            foreach (ExternalConnection conn in workbook.DataConnections)
            {
                if (conn.Name.Equals(targetConnectionName, StringComparison.OrdinalIgnoreCase))
                {
                    targetConnection = conn;
                    break;
                }
            }

            if (targetConnection == null)
            {
                Console.WriteLine($"Connection \"{targetConnectionName}\" not found in the workbook.");
                workbook.Save("OutputWorkbook.xlsx");
                return;
            }

            // Prepare lists to hold found QueryTables and ListObjects with their worksheets
            List<(Worksheet sheet, QueryTable qt)> linkedQueryTables = new List<(Worksheet, QueryTable)>();
            List<(Worksheet sheet, ListObject lo)> linkedListObjects = new List<(Worksheet, ListObject)>();

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Check QueryTables
                foreach (QueryTable qt in sheet.QueryTables)
                {
                    ExternalConnection qtConn = qt.ExternalConnection;
                    if (qtConn != null && qtConn.Id == targetConnection.Id)
                    {
                        linkedQueryTables.Add((sheet, qt));
                    }
                }

                // Check ListObjects
                foreach (ListObject lo in sheet.ListObjects)
                {
                    QueryTable loQt = lo.QueryTable;
                    if (loQt != null)
                    {
                        ExternalConnection loConn = loQt.ExternalConnection;
                        if (loConn != null && loConn.Id == targetConnection.Id)
                        {
                            linkedListObjects.Add((sheet, lo));
                        }
                    }
                }
            }

            // Output results
            Console.WriteLine($"External Connection \"{targetConnectionName}\" (Id: {targetConnection.Id}) is linked to:");
            Console.WriteLine($"- QueryTables: {linkedQueryTables.Count}");
            foreach (var item in linkedQueryTables)
            {
                Console.WriteLine($"  Worksheet: {item.sheet.Name}, QueryTable Name: {item.qt.Name}");
            }

            Console.WriteLine($"- ListObjects: {linkedListObjects.Count}");
            foreach (var item in linkedListObjects)
            {
                Console.WriteLine($"  Worksheet: {item.sheet.Name}, ListObject Name: {item.lo.DisplayName}");
            }

            // Save the workbook (no modifications made, but required by lifecycle rule)
            workbook.Save("OutputWorkbook.xlsx");
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            IdentifyLinkedObjects.Run();
        }
    }
}
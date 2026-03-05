using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;
using Aspose.Cells.Tables;

class FindExternalDataConnections
{
    static void Main()
    {
        // Load the workbook from a file (replace with your actual file path)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Console.WriteLine($"Worksheet: {sheet.Name}");

            // ----------- QueryTables -----------
            if (sheet.QueryTables.Count > 0)
            {
                Console.WriteLine($"  QueryTables count: {sheet.QueryTables.Count}");
                foreach (QueryTable qt in sheet.QueryTables)
                {
                    Console.WriteLine($"    QueryTable Name: {qt.Name}");

                    // Get the external connection associated with the query table
                    ExternalConnection extConn = qt.ExternalConnection;
                    if (extConn != null)
                    {
                        Console.WriteLine($"      Connection Id: {extConn.Id}");
                        Console.WriteLine($"      Connection Name: {extConn.Name}");
                        Console.WriteLine($"      Class Type: {extConn.ClassType}");
                        Console.WriteLine($"      Connection String: {extConn.ConnectionString}");
                    }
                    else
                    {
                        Console.WriteLine("      No external connection associated with this query table.");
                    }
                }
            }

            // ----------- ListObjects (Tables) -----------
            if (sheet.ListObjects.Count > 0)
            {
                Console.WriteLine($"  ListObjects count: {sheet.ListObjects.Count}");
                foreach (ListObject lo in sheet.ListObjects)
                {
                    // Use DisplayName as the name of the ListObject
                    Console.WriteLine($"    ListObject Name: {lo.DisplayName}");

                    // Attempt to retrieve an ExternalConnection property via reflection
                    var extConnProp = lo.GetType().GetProperty("ExternalConnection");
                    if (extConnProp != null)
                    {
                        var extConn = extConnProp.GetValue(lo) as ExternalConnection;
                        if (extConn != null)
                        {
                            Console.WriteLine($"      Connection Id: {extConn.Id}");
                            Console.WriteLine($"      Connection Name: {extConn.Name}");
                            Console.WriteLine($"      Class Type: {extConn.ClassType}");
                        }
                        else
                        {
                            Console.WriteLine("      ListObject has no external connection.");
                        }
                    }
                    else
                    {
                        Console.WriteLine("      ListObject does not expose an ExternalConnection property.");
                    }
                }
            }
        }

        // Save the workbook (unchanged) – optional if you need to persist any modifications
        workbook.Save("output.xlsx");
    }
}
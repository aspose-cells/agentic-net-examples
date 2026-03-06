using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    public class GetODataConnectionInfo
    {
        public static void Run()
        {
            Workbook workbook = new Workbook("input.xlsx");
            ExternalConnectionCollection connections = workbook.DataConnections;

            if (connections.Count == 0)
            {
                Console.WriteLine("No external data connections found in the workbook.");
                return;
            }

            for (int i = 0; i < connections.Count; i++)
            {
                ExternalConnection conn = connections[i];
                if (conn.ClassType == ExternalConnectionClassType.DataModel)
                {
                    Console.WriteLine($"--- OData Connection #{i + 1} ---");
                    Console.WriteLine($"Name               : {conn.Name}");
                    Console.WriteLine($"Class Type         : {conn.ClassType}");
                    Console.WriteLine($"Connection String  : {conn.ConnectionString}");
                    Console.WriteLine($"Command            : {conn.Command}");
                    Console.WriteLine($"Command Type       : {conn.CommandType}");
                    Console.WriteLine($"Source Type        : {conn.SourceType}");
                    Console.WriteLine($"Connection File    : {conn.ConnectionFile}");
                    Console.WriteLine($"Odc File           : {conn.OdcFile}");
                    Console.WriteLine();
                }
            }

            workbook.Save("output.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            GetODataConnectionInfo.Run();
        }
    }
}
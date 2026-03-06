using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class RetrieveSqlConnectionData
{
    static void Main()
    {
        string inputPath = "InputWorkbook.xlsx";

        // Ensure the input file exists; create an empty workbook if it does not.
        if (!File.Exists(inputPath))
        {
            var tempWb = new Workbook();
            tempWb.Save(inputPath);
        }

        Workbook workbook = new Workbook(inputPath);

        ExternalConnectionCollection connections = workbook.DataConnections;

        foreach (ExternalConnection conn in connections)
        {
            if (conn is DBConnection dbConn)
            {
                Console.WriteLine($"Connection Name   : {dbConn.Name}");
                Console.WriteLine($"Connection String : {dbConn.ConnectionString}");
                Console.WriteLine($"Command           : {dbConn.Command}");
                Console.WriteLine($"Command Type      : {dbConn.CommandType}");
                Console.WriteLine("-------------------------------------------");
            }
        }

        workbook.Save("OutputWorkbook.xlsx");
    }
}
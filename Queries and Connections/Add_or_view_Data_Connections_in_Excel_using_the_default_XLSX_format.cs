using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class Program
{
    static void Main()
    {
        // Create a new workbook (default XLSX format)
        Workbook workbook = new Workbook();

        // Add some sample data to make the workbook meaningful
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("ID");
        sheet.Cells["B1"].PutValue("Name");
        sheet.Cells["A2"].PutValue(1);
        sheet.Cells["B2"].PutValue("Alice");
        sheet.Cells["A3"].PutValue(2);
        sheet.Cells["B3"].PutValue("Bob");

        // Create a DataModelConnection instance.
        // The constructor is internal, so we use FormatterServices to obtain an uninitialized object.
        DataModelConnection connection = (DataModelConnection)FormatterServices.GetUninitializedObject(typeof(DataModelConnection));

        // Set essential properties for the connection
        connection.Name = "SampleDataModelConnection";
        connection.Command = "SELECT * FROM SampleTable";
        connection.CommandType = OLEDBCommandType.SqlStatement;
        connection.ConnectionString = "Provider=SQLOLEDB;Data Source=MyServer;Initial Catalog=MyDB;Integrated Security=SSPI";
        connection.SourceType = ConnectionDataSourceType.DataFeedDataModel;

        // Add the connection to the workbook's DataConnections collection
        ((IList<ExternalConnection>)workbook.DataConnections).Add(connection);

        // View (read) all data connections in the workbook
        Console.WriteLine("DataConnections count: " + workbook.DataConnections.Count);
        for (int i = 0; i < workbook.DataConnections.Count; i++)
        {
            ExternalConnection conn = workbook.DataConnections[i];
            Console.WriteLine($"Connection {i + 1}:");
            Console.WriteLine($"  Name       : {conn.Name}");
            Console.WriteLine($"  ClassType  : {conn.ClassType}");
            Console.WriteLine($"  SourceType : {conn.SourceType}");
            Console.WriteLine($"  Command    : {conn.Command}");
        }

        // Save the workbook as an XLSX file
        workbook.Save("DataConnectionsDemo.xlsx");
    }
}
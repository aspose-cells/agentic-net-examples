using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

public static class ConnectionHelper
{
    /// <summary>
    /// Retrieves a dictionary where the key is the connection name and the value is its Command text.
    /// </summary>
    /// <param name="workbook">An initialized Aspose.Cells Workbook instance.</param>
    /// <returns>Dictionary of connection names and their Command properties.</returns>
    public static Dictionary<string, string> GetConnectionCommands(Workbook workbook)
    {
        if (workbook == null)
            throw new ArgumentNullException(nameof(workbook));

        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        ExternalConnectionCollection connections = workbook.DataConnections;

        for (int i = 0; i < connections.Count; i++)
        {
            ExternalConnection conn = connections[i];
            string name = conn.Name;
            if (string.IsNullOrEmpty(name))
                continue;

            string command = conn.Command ?? string.Empty;
            result[name] = command;
        }

        return result;
    }
}

public class Program
{
    public static void Main()
    {
        // Create an empty workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Example: add a dummy data connection (optional, for demonstration)
        // var conn = workbook.DataConnections.Add("SampleConnection", "SELECT * FROM Table1", "OLEDB", "Provider=SQLOLEDB;Data Source=.;Initial Catalog=MyDb;");
        // conn.Command = "SELECT * FROM Table1";

        // Retrieve connection commands
        Dictionary<string, string> commands = ConnectionHelper.GetConnectionCommands(workbook);

        // Output results
        foreach (var kvp in commands)
        {
            Console.WriteLine($"Connection Name: {kvp.Key}, Command: {kvp.Value}");
        }

        // Keep console window open
        Console.WriteLine("Press any key to exit...");
        Console.ReadKey();
    }
}
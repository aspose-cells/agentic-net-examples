using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class UpdateDbConnectionCredentials
{
    static void Main()
    {
        // Path to the existing workbook that contains the DBConnection
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Load the workbook (lifecycle rule: load)
        Workbook workbook = new Workbook(inputPath);

        // Get the collection of external data connections
        ExternalConnectionCollection connections = workbook.DataConnections;

        // Specify the name of the DBConnection to update
        string targetConnectionName = "MyDBConnection";

        // New credentials
        string newUserName = "newUser";
        string newPassword = "newPassword";

        // Find the DBConnection with the specified name
        DBConnection dbConn = null;
        foreach (ExternalConnection conn in connections)
        {
            if (conn is DBConnection db && db.Name.Equals(targetConnectionName, StringComparison.OrdinalIgnoreCase))
            {
                dbConn = db;
                break;
            }
        }

        if (dbConn == null)
        {
            Console.WriteLine($"DBConnection named '{targetConnectionName}' not found.");
            return;
        }

        // Update the connection string with new User ID and Password.
        // Example connection string format:
        // Provider=SQLOLEDB;Data Source=ServerName;Initial Catalog=DatabaseName;User ID=oldUser;Password=oldPass;
        // We'll replace or add the User ID and Password parts.

        string connStr = dbConn.ConnectionString;

        // Helper to replace or add a key=value pair in the connection string
        string SetOrUpdate(string source, string key, string value)
        {
            string[] parts = source.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            bool found = false;
            for (int i = 0; i < parts.Length; i++)
            {
                string[] kv = parts[i].Split(new[] { '=' }, 2);
                if (kv.Length == 2 && kv[0].Trim().Equals(key, StringComparison.OrdinalIgnoreCase))
                {
                    parts[i] = $"{key}={value}";
                    found = true;
                    break;
                }
            }
            if (!found)
            {
                // Append the new key=value pair
                source = source.TrimEnd(';') + $";{key}={value};";
                return source;
            }
            return string.Join(";", parts) + ";";
        }

        connStr = SetOrUpdate(connStr, "User ID", newUserName);
        connStr = SetOrUpdate(connStr, "Password", newPassword);

        // Assign the updated connection string back to the DBConnection
        dbConn.ConnectionString = connStr;

        // Ensure the password is saved within the connection string
        dbConn.SavePassword = true;

        // Save the workbook (lifecycle rule: save)
        workbook.Save(outputPath);

        Console.WriteLine("Credentials updated and workbook saved successfully.");
    }
}
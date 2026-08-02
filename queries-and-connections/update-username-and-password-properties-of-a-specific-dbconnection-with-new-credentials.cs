using System;
using System.IO;
using System.Data.Common;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class UpdateDbConnectionCredentials
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook that contains the external DB connection
            Workbook workbook = new Workbook(inputPath);

            // Access the collection of external connections
            ExternalConnectionCollection connections = workbook.DataConnections;

            // Define the name of the connection you want to modify
            string targetConnectionName = "MyDBConnection";

            // New credentials
            string newUserName = "newUser";
            string newPassword = "newPassword";

            bool updated = false;

            // Iterate through the connections to find the specific DBConnection
            foreach (ExternalConnection conn in connections)
            {
                if (conn is DBConnection dbConn &&
                    dbConn.Name.Equals(targetConnectionName, StringComparison.OrdinalIgnoreCase))
                {
                    // Parse the existing connection string
                    var builder = new DbConnectionStringBuilder
                    {
                        ConnectionString = dbConn.ConnectionString
                    };

                    // Update User ID and Password (common keywords)
                    if (builder.ContainsKey("User ID"))
                        builder["User ID"] = newUserName;
                    else if (builder.ContainsKey("User"))
                        builder["User"] = newUserName;

                    if (builder.ContainsKey("Password"))
                        builder["Password"] = newPassword;
                    else if (builder.ContainsKey("Pwd"))
                        builder["Pwd"] = newPassword;

                    // Assign the updated connection string back to the DBConnection
                    dbConn.ConnectionString = builder.ConnectionString;

                    // Ensure the password is saved within the connection string
                    dbConn.SavePassword = true;

                    Console.WriteLine($"Updated connection '{targetConnectionName}' with new credentials.");
                    updated = true;
                    break;
                }
            }

            if (!updated)
            {
                Console.WriteLine($"Connection '{targetConnectionName}' not found.");
            }

            // Save the workbook with the modified connection settings
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved as '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
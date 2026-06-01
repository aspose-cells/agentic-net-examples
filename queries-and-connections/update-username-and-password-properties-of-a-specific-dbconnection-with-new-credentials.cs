using System;
using System.Data.Common;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class UpdateDbConnectionCredentials
{
    static void Main()
    {
        // Load the workbook that contains the external DB connection
        Workbook workbook = new Workbook("input.xlsx");

        // Access the collection of external connections
        ExternalConnectionCollection connections = workbook.DataConnections;

        // Define the name of the connection you want to update
        string targetConnectionName = "MyDBConnection";

        // New credentials
        string newUserName = "newUser";
        string newPassword = "newPassword";

        // Iterate through the connections to find the specific DBConnection
        foreach (ExternalConnection conn in connections)
        {
            if (conn is DBConnection dbConn && dbConn.Name.Equals(targetConnectionName, StringComparison.OrdinalIgnoreCase))
            {
                // Parse the existing connection string
                DbConnectionStringBuilder builder = new DbConnectionStringBuilder
                {
                    ConnectionString = dbConn.ConnectionString
                };

                // Update User ID and Password entries (keys may vary depending on provider)
                // Common keys: "User ID", "UID", "User", "Password", "PWD"
                if (builder.ContainsKey("User ID")) builder["User ID"] = newUserName;
                else if (builder.ContainsKey("UID")) builder["UID"] = newUserName;
                else if (builder.ContainsKey("User")) builder["User"] = newUserName;

                if (builder.ContainsKey("Password")) builder["Password"] = newPassword;
                else if (builder.ContainsKey("PWD")) builder["PWD"] = newPassword;

                // Assign the modified connection string back to the DBConnection
                dbConn.ConnectionString = builder.ConnectionString;

                // Ensure the password is saved within the connection string
                dbConn.SavePassword = true;

                // Optionally, update the description to reflect the change
                dbConn.ConnectionDescription = $"Updated credentials for {newUserName}";
            }
        }

        // Save the workbook with the updated connection settings
        workbook.Save("output.xlsx");
    }
}
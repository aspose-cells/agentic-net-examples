// Title: Update DBConnection User ID and Password in an Excel workbook using Aspose.Cells for .NET
// Description: Loads an Excel file, finds a DBConnection (by name or first occurrence), modifies its connection string with new "User ID" and "Password" via DbConnectionStringBuilder, enables password persistence with SavePassword, and saves the workbook.
// Keywords: Aspose.Cells | C# | .NET | DBConnection | ExternalConnection | Excel workbook | update credentials | change User ID | change password | ConnectionStringBuilder | SavePassword | modify external DB connection | global | US
// Common Searches: Aspose.Cells change DBConnection user name | update password for Excel external connection .NET | set SavePassword true Aspose.Cells | find DBConnection by name in workbook | modify connection string with Aspose.Cells
// Developer Intent: Replace the username and password of a specific DBConnection in an Excel file and ensure the new credentials are stored.
// Use Cases: Locate a DBConnection named "MyDbConnection" and assign new credentials. | When the named connection is missing, update the first DBConnection found. | Persist the new password by setting DBConnection.SavePassword before saving.
// AI Prompts: Generate C# code with Aspose.Cells that searches for a DBConnection by name and updates its User ID and Password. | Explain how DbConnectionStringBuilder works with Aspose.Cells to edit a DBConnection's connection string. | Provide a step‑by‑step tutorial for changing external database credentials in an Excel workbook using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;
using System.Data.Common;

// Loads an Excel file, finds a DBConnection (by name or first occurrence), modifies its connection string with new "User ID" and "Password" via DbConnectionStringBuilder, enables password persistence with SavePassword, and saves the workbook.
class UpdateDbConnectionCredentials
{
    static void Main()
    {
        // Load the workbook that contains the external DB connection
        Workbook workbook = new Workbook("input.xlsx");

        // Access the collection of external connections
        ExternalConnectionCollection connections = workbook.DataConnections;

        // Identify the DBConnection to update (by name or first occurrence)
        DBConnection dbConn = null;
        foreach (ExternalConnection conn in connections)
        {
            if (conn is DBConnection db)
            {
                // Example: match by connection name; adjust as needed
                if (db.Name == "MyDbConnection")
                {
                    dbConn = db;
                    break;
                }
            }
        }

        // If not found by name, fallback to the first DBConnection
        if (dbConn == null)
        {
            foreach (ExternalConnection conn in connections)
            {
                if (conn is DBConnection db)
                {
                    dbConn = db;
                    break;
                }
            }
        }

        if (dbConn == null)
        {
            Console.WriteLine("No DBConnection found in the workbook.");
            return;
        }

        // Update the connection string with new user name and password
        var builder = new DbConnectionStringBuilder
        {
            ConnectionString = dbConn.ConnectionString
        };

        // Set new credentials
        builder["User ID"] = "newUserName";
        builder["Password"] = "newPassword";

        // Apply the modified connection string back to the DBConnection
        dbConn.ConnectionString = builder.ConnectionString;

        // Ensure the password is saved within the connection string
        dbConn.SavePassword = true;

        // Save the workbook with the updated connection settings
        workbook.Save("output.xlsx");

        Console.WriteLine("Credentials updated and workbook saved successfully.");
    }
}

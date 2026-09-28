// Title: How to programmatically update the User ID and Password of a named DBConnection in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that searches a workbook’s DataConnections collection for a DBConnection with a specific name and returns its object. | Show how to modify the ConnectionString of a found DBConnection to replace or append the User ID and Password parameters while preserving other settings. | Demonstrate saving the workbook after setting the DBConnection’s SavePassword property to true so the new credentials are persisted.
// Common Searches: c# aspose.cells find dbconnection by name in workbook | aspose.cells update connection string user id password | how to preserve existing parameters when changing Excel data connection credentials in .net | set SavePassword flag for external DB connection using Aspose.Cells | replace missing User ID or Password in Aspose.Cells DBConnection string
// Tags: update DBConnection connection string Aspose.Cells | set SavePassword property Aspose.Cells DBConnection | replace User ID parameter in Excel data connection C# | locate DBConnection by name Aspose.Cells | modify external DB connection credentials .NET

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

// The sample loads an Excel file, locates the DBConnection named "MyConnection", updates its ConnectionString with new User ID and Password while keeping other parameters intact, enables SavePassword, and saves the workbook as a new file.
class UpdateDbConnectionCredentials
{
    static void Main()
    {
        // Load the workbook that contains the DB connection
        Workbook workbook = new Workbook("input.xlsx");

        // Name of the DBConnection to update
        string targetConnectionName = "MyConnection";

        // Locate the DBConnection in the workbook's DataConnections collection
        DBConnection dbConnection = null;
        foreach (ExternalConnection conn in workbook.DataConnections)
        {
            if (conn is DBConnection dbConn && dbConn.Name == targetConnectionName)
            {
                dbConnection = dbConn;
                break;
            }
        }

        if (dbConnection != null)
        {
            // New credentials
            string newUserName = "newUser";
            string newPassword = "newPassword";

            // Update the ConnectionString with the new User ID and Password.
            // Preserve existing parts of the string and replace or append credentials.
            string connStr = dbConnection.ConnectionString;

            // Replace or add User ID
            if (connStr.IndexOf("User ID=", StringComparison.OrdinalIgnoreCase) >= 0)
                connStr = Regex.Replace(connStr, "(User ID=)[^;]*", $"$1{newUserName}", RegexOptions.IgnoreCase);
            else
                connStr = connStr.TrimEnd(';') + $";User ID={newUserName}";

            // Replace or add Password
            if (connStr.IndexOf("Password=", StringComparison.OrdinalIgnoreCase) >= 0)
                connStr = Regex.Replace(connStr, "(Password=)[^;]*", $"$1{newPassword}", RegexOptions.IgnoreCase);
            else
                connStr = connStr.TrimEnd(';') + $";Password={newPassword}";

            // Assign the modified connection string back to the DBConnection
            dbConnection.ConnectionString = connStr;

            // Ensure the password is saved as part of the connection string
            dbConnection.SavePassword = true;
        }
        else
        {
            Console.WriteLine($"DBConnection named '{targetConnectionName}' not found.");
        }

        // Save the workbook with the updated connection information
        workbook.Save("output.xlsx");
    }
}

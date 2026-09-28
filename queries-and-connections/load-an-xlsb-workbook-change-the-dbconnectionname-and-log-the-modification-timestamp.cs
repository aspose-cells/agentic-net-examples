// Title: Rename DBConnection objects in an XLSB workbook with a timestamp and record the change using a custom document property (Aspose.Cells for .NET)
// AI Prompts: Load an XLSB file with Aspose.Cells, enumerate its ExternalConnection collection, assign each DBConnection.Name a value that includes the current DateTime, and persist the workbook. | Create or update a custom document property called "DBConnectionModifiedTime" with the current timestamp after renaming DBConnection objects in an XLSB workbook using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# how to change DBConnection name in an XLSB file | Add custom document property with current date in Excel workbook using Aspose.Cells .NET | Programmatically update external data connections in an XLSB workbook with Aspose.Cells | Log modifications of Excel data connections by storing timestamp in custom property C#
// Tags: rename DBConnection Aspose.Cells XLSB | set custom document property timestamp .NET | enumerate data connections Aspose.Cells | modify data connection name C# | persist XLSB changes Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

// The example loads an XLSB workbook, iterates its data connections, updates each DBConnection.Name with a timestamped identifier, adds a custom document property to capture the modification time, and saves the updated workbook using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Paths to the source and destination XLSB files
        string sourcePath = "input.xlsb";
        string destinationPath = "output.xlsb";

        // Load the XLSB workbook
        Workbook workbook = new Workbook(sourcePath);

        // Iterate through all data connections in the workbook
        foreach (ExternalConnection connection in workbook.DataConnections)
        {
            // Check if the connection is a DBConnection
            if (connection is DBConnection dbConnection)
            {
                // Change the DBConnection name (example: include a timestamp)
                dbConnection.Name = "ModifiedConnection_" + DateTime.Now.ToString("yyyyMMddHHmmss");
            }
        }

        // Log the modification timestamp as a custom document property
        workbook.CustomDocumentProperties.Add("DBConnectionModifiedTime", DateTime.Now);

        // Save the modified workbook
        workbook.Save(destinationPath);
    }
}

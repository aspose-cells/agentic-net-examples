// Title: Rename DBConnection in an XLSB workbook and log the change timestamp with Aspose.Cells for .NET
// Description: C# example that loads an XLSB file, iterates through its DataConnections, updates the Name of any DBConnection with a timestamp, adds a custom document property called ModificationTimestamp, and saves the workbook. Demonstrates Aspose.Cells external‑connection handling and audit‑trail creation.
// Keywords: Aspose.Cells XLSB rename DBConnection | C# modify data connection name | add custom document property timestamp | Aspose.Cells external connections example | audit trail workbook Aspose.Cells | DBConnection.Name C# | Aspose.Cells custom properties
// Common Searches: change DBConnection name in XLSB using Aspose.Cells | add modification timestamp to workbook with Aspose.Cells | iterate data connections in XLSB file C# | Aspose.Cells update database connection name | store custom property in XLSB workbook
// Developer Intent: Update the DBConnection name in an XLSB workbook and record the modification time as a custom property.
// Use Cases: Standardize connection names across multiple workbooks before distribution. | Create an immutable audit log of when data connections were altered. | Automate batch processing to embed timestamps for compliance reporting.
// AI Prompts: Write C# code with Aspose.Cells that renames every DBConnection in an XLSX or XLSB file and saves the operation time in a custom document property. | Explain how to read the ModificationTimestamp custom property from a workbook after it has been saved with Aspose.Cells. | Provide a step‑by‑step tutorial for looping through DataConnections in an XLSB workbook, changing DBConnection.Name based on the current date, and adding a timestamp property.

using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

// C# example that loads an XLSB file, iterates through its DataConnections, updates the Name of any DBConnection with a timestamp, adds a custom document property called ModificationTimestamp, and saves the workbook. Demonstrates Aspose.Cells external‑connection handling and audit‑trail creation.
class ModifyDbConnection
{
    static void Main()
    {
        // Load the existing XLSB workbook
        string inputPath = "input.xlsb";
        Workbook workbook = new Workbook(inputPath);

        // Iterate through all data connections and modify DBConnection.Name if present
        foreach (ExternalConnection connection in workbook.DataConnections)
        {
            if (connection is DBConnection dbConn)
            {
                // Set a new name for the DB connection (example uses a timestamp)
                dbConn.Name = "UpdatedConnection_" + DateTime.Now.ToString("yyyyMMddHHmmss");
            }
        }

        // Log the modification timestamp as a custom document property
        workbook.CustomDocumentProperties.Add("ModificationTimestamp", DateTime.Now);

        // Save the modified workbook
        string outputPath = "output.xlsb";
        workbook.Save(outputPath);
    }
}

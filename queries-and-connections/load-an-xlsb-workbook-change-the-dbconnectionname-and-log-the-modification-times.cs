using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class ModifyDbConnection
{
    static void Main()
    {
        // Path to the XLSB workbook
        string filePath = "input.xlsb";

        // Load the workbook (uses the Workbook(string) constructor)
        Workbook workbook = new Workbook(filePath);

        // Iterate through all data connections and modify the name of DBConnection objects
        foreach (ExternalConnection connection in workbook.DataConnections)
        {
            if (connection is DBConnection dbConn)
            {
                // Change the connection name as required
                dbConn.Name = "UpdatedConnectionName";
            }
        }

        // Log the modification timestamp as a custom document property
        string timestampPropertyName = "ModificationTimestamp";
        string timestampValue = DateTime.UtcNow.ToString("o"); // ISO 8601 format

        // If the property already exists, update it; otherwise, add a new one
        if (workbook.CustomDocumentProperties.Contains(timestampPropertyName))
        {
            workbook.CustomDocumentProperties[timestampPropertyName].Value = timestampValue;
        }
        else
        {
            workbook.CustomDocumentProperties.Add(timestampPropertyName, timestampValue);
        }

        // Save the modified workbook (uses the Workbook.Save(string) method)
        workbook.Save(filePath);
    }
}
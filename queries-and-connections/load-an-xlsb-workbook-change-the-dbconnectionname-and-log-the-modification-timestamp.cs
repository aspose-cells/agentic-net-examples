using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class Program
{
    static void Main()
    {
        // Paths to the input and output XLSB files
        string inputPath = "input.xlsb";
        string outputPath = "output.xlsb";

        // Load the XLSB workbook
        Workbook workbook = new Workbook(inputPath);

        // Iterate through all data connections and modify DBConnection names
        foreach (ExternalConnection connection in workbook.DataConnections)
        {
            if (connection is DBConnection dbConn)
            {
                // Append a suffix to indicate modification
                dbConn.Name = dbConn.Name + "_Modified";
            }
        }

        // Log the modification timestamp as a custom document property
        string propertyName = "DBConnectionModifiedTime";
        workbook.CustomDocumentProperties.Add(propertyName, DateTime.Now);

        // Save the modified workbook
        workbook.Save(outputPath);
    }
}
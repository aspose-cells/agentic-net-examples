using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class ExternalConnectionErrorHandlingDemo
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the collection of external connections
        ExternalConnectionCollection connections = workbook.DataConnections;

        // Iterate through each connection and attempt to set an external source file
        foreach (ExternalConnection conn in connections)
        {
            try
            {
                // Define a new source file path (intentionally invalid for demonstration)
                string newPath = @"C:\InvalidPath\ExternalData.xlsx";

                // Validate the path before assigning; throw if it does not exist
                if (!File.Exists(newPath))
                {
                    throw new FileNotFoundException("External source file not found.", newPath);
                }

                // Assign the valid path to the connection
                conn.SourceFile = newPath;
                Console.WriteLine($"Connection '{conn.Name}' source file set to: {conn.SourceFile}");
            }
            catch (Exception ex)
            {
                // Capture and display any errors related to the invalid path
                Console.WriteLine($"Error updating connection '{conn.Name}': {ex.Message}");
            }
        }

        // Save the workbook after processing
        workbook.Save("output.xlsx");
    }
}
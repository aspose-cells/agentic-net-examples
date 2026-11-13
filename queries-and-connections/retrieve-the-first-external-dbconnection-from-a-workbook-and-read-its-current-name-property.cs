// Title: Read the Name of the First External DBConnection in an Excel Workbook with Aspose.Cells for .NET
// Common Searches: how to obtain the name of the first external data connection in an Excel file using Aspose.Cells C# | using reflection to access ExternalDataConnections collection in an Aspose.Cells worksheet | retrieve external DBConnection name from a workbook when the ExternalDataConnections property is unavailable | C# code to list external data connections in an .xlsx file with Aspose.Cells | handle missing ExternalDataConnections property across different Aspose.Cells versions
// Tags: read external data connection name Aspose.Cells | external connections reflection Aspose.Cells | enumerate workbook external connections .NET | missing ExternalDataConnections handling Aspose.Cells | retrieve DBConnection Name property C#

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, accesses its first worksheet, and uses reflection to safely obtain the ExternalDataConnections collection (if present). It enumerates the connections, reads the Name property of the first DBConnection, and outputs the name, while gracefully handling cases where the property is absent, no connections exist, or runtime errors occur.
class Program
{
    static void Main()
    {
        const string filePath = "input.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(filePath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Attempt to retrieve external data connections if the API is available
            // Note: In some versions of Aspose.Cells the ExternalDataConnections property may not exist.
            // The following code uses reflection to safely access it without compile‑time dependency.
            var connectionsProp = typeof(Worksheet).GetProperty("ExternalDataConnections");
            if (connectionsProp != null)
            {
                var connections = connectionsProp.GetValue(sheet, null) as System.Collections.IEnumerable;
                if (connections != null)
                {
                    var enumerator = connections.GetEnumerator();
                    if (enumerator.MoveNext())
                    {
                        // Use reflection to get the Name property of the first connection
                        var firstConnection = enumerator.Current;
                        var nameProp = firstConnection.GetType().GetProperty("Name");
                        string name = nameProp != null ? nameProp.GetValue(firstConnection, null)?.ToString() : "Unnamed";
                        Console.WriteLine("First external data connection name: " + name);
                    }
                    else
                    {
                        Console.WriteLine("No external data connections found in the workbook.");
                    }
                }
                else
                {
                    Console.WriteLine("ExternalDataConnections property is present but returned null.");
                }
            }
            else
            {
                Console.WriteLine("External data connections are not supported in this version of Aspose.Cells.");
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime errors (e.g., file format issues, loading problems)
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}

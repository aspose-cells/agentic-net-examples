// Title: Export external connection names from an Excel workbook to a plain‑text file with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, iterates all worksheets, extracts each external connection's Name, and saves the list to a .txt file. | Show how to use reflection to access the Worksheet.ExternalConnections collection for compatibility with older Aspose.Cells versions and write the connection names to a text file. | Create a robust utility that checks for the workbook's existence, handles missing ExternalConnections, and outputs the connection names line by line to a specified output path.
// Common Searches: aspacells get list of external data connections from workbook c# | how to write excel connection names to a text file using Aspose.Cells | c# export workbook external connections with reflection for older Aspose.Cells versions | sample code to enumerate external connections in each worksheet Aspose.Cells .NET | save external connection names from input.xlsx to ExternalConnections.txt
// Tags: Aspose.Cells enumerate external connections | C# write connection names to text file | reflection access Worksheet.ExternalConnections | handle missing ExternalConnections in Aspose.Cells | export workbook data connection list .NET

using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Aspose.Cells;

// Loads an Excel workbook with Aspose.Cells, uses reflection to retrieve the ExternalConnections collection from each worksheet (compatible with older library versions), extracts each connection's Name, and writes the names to a plain‑text file, with error handling for missing files and unsupported versions.
class ExportExternalConnections
{
    static void Main()
    {
        // Path to the source workbook
        string workbookPath = "input.xlsx";

        // Path to the output text file
        string outputPath = "ExternalConnections.txt";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: The file \"{workbookPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Prepare a list to hold connection names
            List<string> connectionNames = new List<string>();

            // Use reflection to access ExternalConnections (may not exist in older versions)
            PropertyInfo externalConnProp = typeof(Worksheet).GetProperty("ExternalConnections");

            if (externalConnProp != null)
            {
                // Iterate through each worksheet and collect connection names
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    var connections = externalConnProp.GetValue(sheet) as System.Collections.IEnumerable;
                    if (connections == null) continue;

                    foreach (object conn in connections)
                    {
                        // Retrieve the Name property via reflection
                        PropertyInfo nameProp = conn.GetType().GetProperty("Name");
                        if (nameProp != null)
                        {
                            string name = nameProp.GetValue(conn) as string;
                            if (!string.IsNullOrEmpty(name))
                                connectionNames.Add(name);
                        }
                    }
                }
            }
            else
            {
                Console.WriteLine("External connections are not supported in this version of Aspose.Cells.");
            }

            // Write each connection name to the text file, one per line
            using (StreamWriter writer = new StreamWriter(outputPath))
            {
                foreach (string name in connectionNames)
                {
                    writer.WriteLine(name);
                }
            }

            Console.WriteLine("External connection names have been exported to: " + outputPath);
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors (e.g., loading issues, permission problems)
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}

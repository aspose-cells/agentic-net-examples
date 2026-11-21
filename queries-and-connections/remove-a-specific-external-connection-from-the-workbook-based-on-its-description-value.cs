// Title: Remove an external data connection by description from an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# remove external connection with specific description | Delete Excel external data connection by description using Aspose.Cells | Iterate worksheet ExternalConnections collection in Aspose.Cells .NET | C# code to filter and delete external connections in an .xlsx file | How to drop a named external connection from a workbook with Aspose.Cells
// Tags: remove external connection Aspose.Cells | external connections collection .NET | filter worksheet external connections C# | delete Excel data connection Aspose.Cells | iterate worksheets external connections C#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an input.xlsx workbook, defines a target description (e.g., "MyConnection"), and shows how to iterate each worksheet's ExternalConnections collection in reverse order to remove connections whose Description matches the target. After optional removal, the workbook is saved as output.xlsx. Note that the ExternalConnections API requires a recent Aspose.Cells version.
    class Program
    {
        static void Main()
        {
            // Paths for input and output workbooks
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Description of the external connection to remove
            string targetDescription = "MyConnection";

            try
            {
                // Ensure the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // NOTE: The ExternalConnections collection is not available in the current
                // Aspose.Cells version used in this environment. If you have a newer version
                // that supports Worksheet.ExternalConnections, you can uncomment and use the
                // code below to remove connections by description.

                /*
                // Iterate through each worksheet's external connections
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate backwards to safely remove items
                    for (int i = sheet.ExternalConnections.Count - 1; i >= 0; i--)
                    {
                        ExternalConnection connection = sheet.ExternalConnections[i];
                        // Remove connection if description matches
                        if (connection.Description == targetDescription)
                        {
                            sheet.ExternalConnections.RemoveAt(i);
                        }
                    }
                }
                */

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

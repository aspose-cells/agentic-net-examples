// Title: How to read the external data connection of each PivotTable in an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx workbook with Aspose.Cells, iterates over all PivotTables, and prints the name of the external data connection for each table, handling cases where the connection information is unavailable. | Create a reusable C# method that accepts a Workbook object and returns a dictionary mapping PivotTable names to their ConnectionName or ExternalConnection, using reflection when the property is not directly exposed. | Generate a console application example that logs pivot table connection details, includes robust error handling for missing properties, and optionally saves the workbook after processing.
// Common Searches: Aspose.Cells C# get external connection name from PivotTable | Read data source of a PivotTable using Aspose.Cells .NET | List all PivotTable connections in an Excel workbook with Aspose.Cells | C# enumerate PivotTables and retrieve their ConnectionName property via Aspose.Cells | Aspose.Cells reflection get PivotTable ConnectionName when not exposed
// Tags: Aspose.Cells read pivot table data connection | C# enumerate pivot tables Aspose.Cells | retrieve pivot table source property .NET | handle missing ConnectionName property Aspose.Cells | extract pivot table connection info Excel C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Namespace containing PivotTable

// The example loads an Excel workbook, accesses each worksheet's PivotTables, and attempts to read the associated external data connection name using the ConnectionName property (via reflection when necessary). It logs the connection details to the console and includes error handling for absent connection information.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure the workbook contains at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("The workbook does not contain any worksheets.");
                return;
            }

            // Access the first worksheet (or any worksheet that contains the pivot table)
            Worksheet sheet = workbook.Worksheets[0];

            // Loop through all pivot tables in the worksheet
            foreach (PivotTable pivot in sheet.PivotTables)
            {
                Console.WriteLine($"Pivot Table: {pivot.Name}");

                // Attempt to display the associated connection name if available
                // (PivotTable may expose a ConnectionName property in some versions)
                try
                {
                    var connectionNameProp = pivot.GetType().GetProperty("ConnectionName");
                    if (connectionNameProp != null)
                    {
                        string connName = connectionNameProp.GetValue(pivot) as string;
                        if (!string.IsNullOrEmpty(connName))
                        {
                            Console.WriteLine($"Associated Connection: {connName}");
                        }
                        else
                        {
                            Console.WriteLine("No associated external connection.");
                        }
                    }
                    else
                    {
                        Console.WriteLine("Connection information not available in this API version.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error retrieving connection info: {ex.Message}");
                }

                Console.WriteLine();
            }

            // Save the workbook if any modifications were made (optional)
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

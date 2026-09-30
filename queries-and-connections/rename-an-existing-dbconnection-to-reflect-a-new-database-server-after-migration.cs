// Title: Rename a DbConnection in an Excel workbook after server migration using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an Excel file with Aspose.Cells, locates a DbConnection named 'OldServerConnection' in the workbook's DbConnectionCollection (using reflection if necessary), renames it to 'NewServerConnection', and saves the updated workbook. | Create a method that, after renaming a DbConnection in an Aspose.Cells workbook, updates its ConnectionString to point to the new database server, handling cases where the DbConnectionCollection property is not directly exposed. | Show how to safely detect the DbConnectionCollection property via reflection, rename the matching connection, and prevent runtime exceptions while working with Aspose.Cells for .NET.
// Common Searches: how to change the name of a data connection in an Excel file with Aspose.Cells C# | rename DbConnection in Aspose.Cells workbook after moving database server | update Excel workbook DB connection string using Aspose.Cells .NET reflection | Aspose.Cells C# find and rename existing DB connection in workbook | access DbConnectionCollection in Aspose.Cells when property is hidden
// Tags: Aspose.Cells rename DbConnection | C# update Excel data connection name | reflection access DbConnectionCollection Aspose.Cells | modify connection string after server migration .NET | Aspose.Cells workbook DB connection handling

using System;
using System.IO;
using System.Linq;
using System.Collections;
using Aspose.Cells;

// Loads an Excel workbook, uses reflection to locate the DbConnectionCollection, finds the connection named 'OldServerConnection', renames it to 'NewServerConnection', optionally updates its connection string, and saves the workbook with the new connection name.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "InputWorkbook.xlsx";
            const string outputPath = "OutputWorkbook.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook that contains the existing DB connection
            var workbook = new Workbook(inputPath);

            // Define the old connection name and the new name you want to assign
            string oldConnectionName = "OldServerConnection";
            string newConnectionName = "NewServerConnection";

            // Attempt to access the DbConnectionCollection via reflection (API may differ between versions)
            var dbConnCollectionProp = workbook.GetType().GetProperty("DbConnectionCollection");
            if (dbConnCollectionProp != null)
            {
                var dbConnections = dbConnCollectionProp.GetValue(workbook) as IEnumerable;
                if (dbConnections != null)
                {
                    foreach (var connection in dbConnections)
                    {
                        var nameProp = connection.GetType().GetProperty("Name");
                        if (nameProp != null && string.Equals(nameProp.GetValue(connection) as string, oldConnectionName, StringComparison.OrdinalIgnoreCase))
                        {
                            // Rename the connection
                            nameProp.SetValue(connection, newConnectionName);

                            // Optional: update the connection string (example for OLE DB)
                            // var connStrProp = connection.GetType().GetProperty("ConnectionString");
                            // if (connStrProp != null)
                            // {
                            //     string oldConnStr = connStrProp.GetValue(connection) as string;
                            //     string newConnStr = oldConnStr?.Replace("Data Source=OldServer;", "Data Source=NewServer;");
                            //     connStrProp.SetValue(connection, newConnStr);
                            // }

                            Console.WriteLine($"Connection renamed to '{newConnectionName}'.");
                            break;
                        }
                    }
                }
                else
                {
                    Console.WriteLine("DbConnectionCollection is empty or not accessible.");
                }
            }
            else
            {
                Console.WriteLine("DbConnectionCollection property not found in this Aspose.Cells version.");
            }

            // Save the workbook with the updated connection name
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved as '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

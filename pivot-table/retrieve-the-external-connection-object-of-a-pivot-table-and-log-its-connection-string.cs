// Title: How to retrieve a PivotTable's external connection string in C# using Aspose.Cells
// Common Searches: asp.net get external connection string from Excel pivot table using Aspose.Cells | c# retrieve external data source of a pivot table with Aspose.Cells | how to use reflection to read pivot table connection string in Aspose.Cells | read external connection of Excel pivot table in C# Aspose.Cells example | Aspose.Cells code to extract pivot table external connection string
// Tags: Aspose.Cells pivot external connection | C# reflection external data source | Excel pivot external connection string | Aspose.Cells get pivot data source | Workbook pivot connection retrieval

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable related classes

// The sample loads an Excel workbook, checks for a PivotTable, and uses reflection to access the internal 'ExternalDataSource' property of the PivotTable. It then navigates to the 'Connection' object and reads its 'ConnectionString', outputting the value to the console while handling missing properties and errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook containing the pivot table
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Retrieve the first pivot table
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Attempt to obtain external connection information via reflection
            try
            {
                PropertyInfo externalProp = typeof(PivotTable).GetProperty("ExternalDataSource", BindingFlags.Public | BindingFlags.Instance);
                if (externalProp != null)
                {
                    object externalDataSource = externalProp.GetValue(pivotTable);
                    if (externalDataSource != null)
                    {
                        // Try to get the Connection property from the external data source object
                        PropertyInfo connectionProp = externalDataSource.GetType().GetProperty("Connection", BindingFlags.Public | BindingFlags.Instance);
                        if (connectionProp != null)
                        {
                            object connectionObj = connectionProp.GetValue(externalDataSource);
                            if (connectionObj != null)
                            {
                                // Finally, get the ConnectionString property
                                PropertyInfo connStrProp = connectionObj.GetType().GetProperty("ConnectionString", BindingFlags.Public | BindingFlags.Instance);
                                if (connStrProp != null)
                                {
                                    string connStr = connStrProp.GetValue(connectionObj) as string;
                                    Console.WriteLine("External Connection String: " + connStr);
                                }
                                else
                                {
                                    Console.WriteLine("ConnectionString property not found.");
                                }
                            }
                            else
                            {
                                Console.WriteLine("Connection object is null.");
                            }
                        }
                        else
                        {
                            Console.WriteLine("Connection property not found on external data source.");
                        }
                    }
                    else
                    {
                        Console.WriteLine("The pivot table is not linked to an external connection.");
                    }
                }
                else
                {
                    Console.WriteLine("External data source information is not supported in this Aspose.Cells version.");
                }
            }
            catch (Exception ex)
            {
                // Any reflection-related errors are caught here
                Console.WriteLine("Error while retrieving external data source: " + ex.Message);
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}

// Title: Read external connection information of a PivotTable using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# get external data source details from a PivotTable | How to extract connection string of a PivotTable with Aspose.Cells .NET | Retrieve PivotTable external connection name and provider using Aspose.Cells | Using reflection to access PivotTable ExternalDataSource in C# Aspose.Cells | Example code for reading PivotTable external connection properties in .NET
// Tags: Aspose.Cells PivotTable external data source | C# read PivotTable connection string | Aspose.Cells retrieve external connection details | Reflection access PivotTable ExternalDataSource | Excel workbook load PivotTable external connection .NET

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable class

namespace PivotTableExternalConnectionDemo
{
    // The sample loads an Excel workbook, selects the first worksheet and its first PivotTable, then uses reflection to obtain the PivotTable's ExternalDataSource object. It reads the connection name, connection string, provider, command text, and command type, and prints these external connection details to the console while handling missing properties and errors gracefully.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (or specify by name/index)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one PivotTable
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No PivotTables found in the worksheet.");
                    return;
                }

                // Get the first PivotTable
                PivotTable pivotTable = worksheet.PivotTables[0];

                // Use reflection to obtain the external data source (avoids compile‑time dependency on a missing type)
                PropertyInfo externalDataSourceProp = pivotTable.GetType().GetProperty("ExternalDataSource", BindingFlags.Public | BindingFlags.Instance);
                if (externalDataSourceProp == null)
                {
                    Console.WriteLine("The PivotTable does not expose an ExternalDataSource property in this Aspose.Cells version.");
                    return;
                }

                object externalDataSource = externalDataSourceProp.GetValue(pivotTable);
                if (externalDataSource == null)
                {
                    Console.WriteLine("The PivotTable does not use an external data connection.");
                    return;
                }

                // Retrieve connection details via reflection
                string connectionName = GetPropertyValue<string>(externalDataSource, "ConnectionName");
                string connectionString = GetPropertyValue<string>(externalDataSource, "ConnectionString");
                string provider = GetPropertyValue<string>(externalDataSource, "Provider");
                string command = GetPropertyValue<string>(externalDataSource, "Command");
                object commandTypeObj = GetPropertyValue<object>(externalDataSource, "CommandType");
                string commandType = commandTypeObj?.ToString() ?? "N/A";

                // Output the retrieved information
                Console.WriteLine("External Connection Details:");
                Console.WriteLine($"Name              : {connectionName}");
                Console.WriteLine($"Connection String : {connectionString}");
                Console.WriteLine($"Provider          : {provider}");
                Console.WriteLine($"Command           : {command}");
                Console.WriteLine($"Command Type      : {commandType}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Helper method to safely get property values via reflection
        private static T GetPropertyValue<T>(object obj, string propertyName)
        {
            try
            {
                PropertyInfo prop = obj.GetType().GetProperty(propertyName, BindingFlags.Public | BindingFlags.Instance);
                if (prop != null && prop.GetValue(obj) is T value)
                {
                    return value;
                }
            }
            catch
            {
                // Ignored – return default if any issue occurs
            }
            return default;
        }
    }
}

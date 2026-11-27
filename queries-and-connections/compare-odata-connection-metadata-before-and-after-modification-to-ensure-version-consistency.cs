// Title: Compare OData connection metadata before and after modification and ensure ODataVersion stays consistent using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells compare OData connection properties before and after editing | verify ODataVersion unchanged after updating CommandText in an Excel workbook with Aspose.Cells | use reflection to read OData connection metadata in Aspose.Cells .NET | snapshot external connection details in Aspose.Cells and detect version changes | how to programmatically modify OData CommandText while keeping version stable in C#
// Tags: Aspose.Cells OData connection metadata extraction | C# reflection external connection properties | compare ODataVersion consistency after CommandText update | programmatic OData CommandText modification in Excel | verify OData connection version stability with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace ODataConnectionMetadataComparer
{
    // The sample loads or creates an Excel workbook, adds an OData connection if needed, captures its Name, Url, ODataVersion, and CommandText via reflection into a ConnectionMetadata object, updates the CommandText property, captures the metadata again, compares the ODataVersion values to confirm they remain unchanged, and saves the modified workbook.
    class Program
    {
        // Helper class to hold a snapshot of connection metadata
        private class ConnectionMetadata
        {
            public string Name { get; set; } = string.Empty;
            public string Url { get; set; } = string.Empty;
            public string ODataVersion { get; set; } = string.Empty;
            public string CommandText { get; set; } = string.Empty;

            // Build metadata from any ExternalConnection using reflection
            public static ConnectionMetadata FromExternalConnection(ExternalConnection conn)
            {
                var meta = new ConnectionMetadata
                {
                    Name = conn.Name,
                    Url = (GetPropertyValue(conn, "Url") as string) ?? "N/A",
                    ODataVersion = GetPropertyValue(conn, "ODataVersion")?.ToString() ?? "None",
                    CommandText = (GetPropertyValue(conn, "CommandText") as string) ?? string.Empty
                };
                return meta;
            }

            // Helper to read property via reflection safely
            private static object? GetPropertyValue(object obj, string propertyName)
            {
                try
                {
                    var prop = obj.GetType().GetProperty(propertyName);
                    return prop?.GetValue(obj);
                }
                catch
                {
                    return null;
                }
            }

            public override string ToString()
            {
                return $"Name: {Name}, Url: {Url}, ODataVersion: {ODataVersion}, CommandText: {CommandText}";
            }
        }

        static void Main(string[] args)
        {
            const string sourcePath = "source.xlsx";
            const string modifiedPath = "modified.xlsx";

            // Ensure Aspose.Cells license is set if you have one.
            // License license = new License();
            // license.SetLicense("Aspose.Cells.lic");

            Workbook workbook = null;

            try
            {
                // Load existing workbook if it exists
                if (File.Exists(sourcePath))
                {
                    workbook = new Workbook(sourcePath);
                    Console.WriteLine($"Loaded workbook from \"{sourcePath}\".");
                }
                else
                {
                    // Create a new workbook and add a dummy OData connection
                    workbook = new Workbook();
                    Console.WriteLine("Source workbook not found. Created a new workbook.");

                    string connName = "DemoOData";
                    string serviceUrl = "https://services.odata.org/V4/Northwind/Northwind.svc/Products";

                    // Add OData connection via reflection to avoid compile‑time dependency on the method signature
                    try
                    {
                        var addMethod = workbook.DataConnections.GetType().GetMethod("AddODataConnection");
                        if (addMethod != null)
                        {
                            addMethod.Invoke(workbook.DataConnections, new object[] { connName, serviceUrl, string.Empty });
                            Console.WriteLine($"Added OData connection \"{connName}\" with URL \"{serviceUrl}\".");
                        }
                        else
                        {
                            Console.WriteLine("AddODataConnection method not available in this Aspose.Cells version.");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error adding OData connection: {ex.Message}");
                    }

                    // Save the initial workbook
                    try
                    {
                        workbook.Save(sourcePath);
                        Console.WriteLine($"Saved initial workbook to \"{sourcePath}\".");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error saving initial workbook: {ex.Message}");
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading or creating workbook: {ex.Message}");
                return;
            }

            // Locate the first OData connection (identified by its runtime type name)
            ExternalConnection odataConnection = null;
            foreach (ExternalConnection extConn in workbook.DataConnections)
            {
                if (extConn.GetType().Name.Equals("ODataConnection", StringComparison.OrdinalIgnoreCase))
                {
                    odataConnection = extConn;
                    break;
                }
            }

            if (odataConnection == null)
            {
                Console.WriteLine("No OData connection found in the workbook. Exiting.");
                return;
            }

            // Capture metadata BEFORE modification
            ConnectionMetadata beforeMetadata = ConnectionMetadata.FromExternalConnection(odataConnection);
            Console.WriteLine("\n--- Metadata BEFORE modification ---");
            Console.WriteLine(beforeMetadata);

            // Modify CommandText if the property exists
            try
            {
                var cmdProp = odataConnection.GetType().GetProperty("CommandText");
                if (cmdProp != null && cmdProp.CanWrite)
                {
                    string originalCommand = cmdProp.GetValue(odataConnection) as string ?? string.Empty;
                    string modifiedCommand = "$select=ProductID,ProductName,UnitPrice";
                    cmdProp.SetValue(odataConnection, modifiedCommand);
                    Console.WriteLine($"\nModified CommandText from \"{originalCommand}\" to \"{modifiedCommand}\".");
                }
                else
                {
                    Console.WriteLine("\nCommandText property not available on this connection type.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\nError modifying CommandText: {ex.Message}");
            }

            // Capture metadata AFTER modification
            ConnectionMetadata afterMetadata = ConnectionMetadata.FromExternalConnection(odataConnection);
            Console.WriteLine("\n--- Metadata AFTER modification ---");
            Console.WriteLine(afterMetadata);

            // Compare version consistency
            bool versionConsistent = string.Equals(beforeMetadata.ODataVersion, afterMetadata.ODataVersion,
                StringComparison.OrdinalIgnoreCase);
            Console.WriteLine("\n--- Version Consistency Check ---");
            if (versionConsistent)
            {
                Console.WriteLine("SUCCESS: OData version remained consistent after modification.");
            }
            else
            {
                Console.WriteLine("FAILURE: OData version changed after modification.");
                Console.WriteLine($"Before: {beforeMetadata.ODataVersion}");
                Console.WriteLine($"After : {afterMetadata.ODataVersion}");
            }

            // Save the workbook with modifications
            try
            {
                workbook.Save(modifiedPath);
                Console.WriteLine($"\nModified workbook saved to \"{modifiedPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\nError saving modified workbook: {ex.Message}");
            }

            Console.WriteLine("\nPress any key to exit...");
            Console.ReadKey();
        }
    }
}

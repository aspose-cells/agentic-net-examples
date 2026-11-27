// Title: Programmatically refresh a named external SQL connection in an Excel workbook using Aspose.Cells for .NET
// Common Searches: how to refresh a specific external SQL connection in an Excel file using Aspose.Cells C# | Aspose.Cells .NET refresh external data connection by name | disable background refresh and call RefreshData on ExternalConnection with Aspose.Cells | fallback to Workbook.RefreshDataConnections when ExternalConnection.RefreshData not available | command line tool to update external SQL connection in XLSX using Aspose.Cells
// Tags: Aspose.Cells sql connection programmatic refresh | ExternalConnection.RefreshData method | Workbook.RefreshDataConnections alternative | background refresh false Aspose.Cells | cli refresh excel connection Aspose.Cells

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace RefreshExternalSqlConnection
{
    // Loads an XLSX workbook, selects a named (or first) external SQL connection, sets BackgroundRefresh to false, attempts to refresh the connection via ExternalConnection.RefreshData using reflection, falls back to Workbook.RefreshDataConnections if needed, and saves the updated workbook to the specified output path.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Validate arguments
                if (args.Length < 2)
                {
                    Console.WriteLine("Insufficient arguments.");
                    Console.WriteLine("Usage: RefreshExternalSqlConnection.exe <inputWorkbookPath> <outputWorkbookPath> [connectionName]");
                    return;
                }

                string inputPath = args[0];
                string outputPath = args[1];
                string targetConnectionName = args.Length >= 3 ? args[2] : null;

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
                    return;
                }

                // Load the workbook
                Console.WriteLine($"Loading workbook \"{inputPath}\"...");
                Workbook workbook = new Workbook(inputPath);

                // Ensure the workbook contains external connections
                if (workbook.DataConnections == null || workbook.DataConnections.Count == 0)
                {
                    Console.WriteLine("The workbook does not contain any external data connections.");
                    return;
                }

                // Locate the desired connection
                ExternalConnection connection = null;
                if (!string.IsNullOrEmpty(targetConnectionName))
                {
                    foreach (ExternalConnection conn in workbook.DataConnections)
                    {
                        if (string.Equals(conn.Name, targetConnectionName, StringComparison.OrdinalIgnoreCase))
                        {
                            connection = conn;
                            break;
                        }
                    }

                    if (connection == null)
                    {
                        Console.WriteLine($"Connection named \"{targetConnectionName}\" was not found.");
                        return;
                    }
                }
                else
                {
                    // Use the first connection if no name is supplied
                    connection = workbook.DataConnections[0];
                    Console.WriteLine($"No connection name supplied – using first connection \"{connection.Name}\".");
                }

                // Ensure refresh happens synchronously (foreground)
                connection.BackgroundRefresh = false;

                // Attempt to refresh the connection using reflection (covers multiple API versions)
                bool refreshed = false;
                try
                {
                    MethodInfo connRefresh = connection.GetType().GetMethod("RefreshData", BindingFlags.Public | BindingFlags.Instance);
                    if (connRefresh != null)
                    {
                        Console.WriteLine($"Refreshing connection \"{connection.Name}\" via ExternalConnection.RefreshData()...");
                        connRefresh.Invoke(connection, null);
                        refreshed = true;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error invoking ExternalConnection.RefreshData(): {ex.Message}");
                }

                // Fallback: try workbook-wide refresh if connection-specific refresh was unavailable
                if (!refreshed)
                {
                    try
                    {
                        MethodInfo wbRefresh = workbook.GetType().GetMethod("RefreshDataConnections", BindingFlags.Public | BindingFlags.Instance);
                        if (wbRefresh != null)
                        {
                            Console.WriteLine("Refreshing all data connections via Workbook.RefreshDataConnections()...");
                            wbRefresh.Invoke(workbook, null);
                            refreshed = true;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error invoking Workbook.RefreshDataConnections(): {ex.Message}");
                    }
                }

                if (!refreshed)
                {
                    Console.WriteLine("Refresh operation is not supported in the current Aspose.Cells version.");
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the updated workbook
                Console.WriteLine($"Saving workbook to \"{outputPath}\"...");
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine("Operation completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}

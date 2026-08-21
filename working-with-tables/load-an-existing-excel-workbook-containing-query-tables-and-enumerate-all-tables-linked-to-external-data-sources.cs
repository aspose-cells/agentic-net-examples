// Title: C# – Enumerate Excel Query Tables and Their External Connections with Aspose.Cells
// Description: Loads an Excel workbook, scans every worksheet for query tables, and prints each table’s name, result range, and linked external‑connection details (ID, name, type, connection string, refresh setting). Also lists workbook‑level external links and saves the enumeration to a new file.
// Keywords: Aspose.Cells query tables enumeration | list external connections C# | Excel query table details Aspose | external links workbook Aspose.Cells | C# enumerate QueryTable | Aspose.Cells ExternalConnection | Excel data source audit .NET
// Common Searches: how to list query tables in an Excel file using Aspose.Cells | retrieve external connection info for Excel query tables C# | get workbook level external links Aspose.Cells | display result range of query tables Aspose.Cells .NET | enumerate query tables and connections programmatically
// Developer Intent: The developer needs to load an existing workbook, identify all query tables, extract their external‑connection properties, and optionally capture workbook‑wide external links for reporting or validation.
// Use Cases: Create an audit report of every query table and its data source before publishing a workbook. | Validate that all query tables have proper external connections and refresh settings. | Extract and update workbook‑level external link definitions in bulk.
// AI Prompts: Generate C# code with Aspose.Cells that iterates through all worksheets, enumerates query tables, and prints each table’s name, result range, and external connection details. | Write a method that returns a list of ExternalLink objects from a workbook and logs their data source strings. | Provide a reusable function that accepts a workbook path, enumerates query tables with connection info, and writes the results to a new Excel file.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsQueryTableEnumeration
{
    // Loads an Excel workbook, scans every worksheet for query tables, and prints each table’s name, result range, and linked external‑connection details (ID, name, type, connection string, refresh setting). Also lists workbook‑level external links and saves the enumeration to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook that should contain query tables
            string workbookPath = "InputWorkbookWithQueryTables.xlsx";

            try
            {
                // Ensure the input file exists; create a placeholder if it does not
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Input file '{workbookPath}' not found. Creating an empty workbook as a placeholder.");
                    Workbook placeholder = new Workbook();
                    placeholder.Worksheets[0].Name = "PlaceholderSheet";
                    placeholder.Save(workbookPath);
                }

                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    if (sheet.QueryTables.Count > 0)
                    {
                        Console.WriteLine($"Worksheet: {sheet.Name}");
                        for (int i = 0; i < sheet.QueryTables.Count; i++)
                        {
                            QueryTable queryTable = sheet.QueryTables[i];
                            Console.WriteLine($"  Query Table {i + 1}: {queryTable.Name}");
                            Console.WriteLine($"    Result Range: {queryTable.ResultRange.Address}");

                            // External connection details
                            ExternalConnection extConn = queryTable.ExternalConnection;
                            if (extConn != null)
                            {
                                Console.WriteLine("    Linked to external data source:");
                                Console.WriteLine($"      Connection ID   : {extConn.Id}");
                                Console.WriteLine($"      Connection Name : {extConn.Name}");
                                Console.WriteLine($"      Connection Type : {extConn.ClassType}");
                                Console.WriteLine($"      Connection String: {extConn.ConnectionString}");
                                Console.WriteLine($"      Refresh on Load : {extConn.RefreshOnLoad}");
                            }
                            else
                            {
                                Console.WriteLine("    No external connection associated with this query table.");
                            }
                        }
                    }
                }

                // List workbook‑level external links
                ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;
                if (externalLinks.Count > 0)
                {
                    Console.WriteLine("\nWorkbook-level external links:");
                    for (int i = 0; i < externalLinks.Count; i++)
                    {
                        ExternalLink link = externalLinks[i];
                        Console.WriteLine($"  External Link {i + 1}: DataSource = {link.DataSource}");
                    }
                }
                else
                {
                    Console.WriteLine("\nNo workbook-level external links found.");
                }

                // Save the processed workbook
                workbook.Save("EnumeratedQueryTablesOutput.xlsx");
                Console.WriteLine("\nProcessing completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

// Title: Refresh All External Data Connections in an Excel Workbook with Aspose.Cells for .NET
// Description: Load an existing workbook, enable RefreshOnLoad for each external data connection, recalculate formulas, and save the updated file to guarantee data consistency across the workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells refresh external connections | C# RefreshOnLoad property | update Excel data connections programmatically | recalculate formulas after data refresh | save workbook after connection refresh | Aspose.Cells ExternalConnectionCollection | automated Excel reporting .NET
// Common Searches: how to refresh all external connections in Excel using Aspose.Cells C# | set RefreshOnLoad for data connections Aspose.Cells | recalculate formulas after refreshing external data connections .NET | save workbook after updating external connections Aspose.Cells | programmatically refresh ODBC or web query connections in Excel
// Developer Intent: Programmatically enable refresh for every external data connection, recalculate dependent formulas, and write the refreshed workbook to a new file.
// Use Cases: Automated reporting pipelines that must pull the latest data from web queries, ODBC, or other external sources before generating final Excel files. | Scheduled tasks that process multiple workbooks, ensuring each file’s connections are refreshed and formulas are up‑to‑date. | Integration of Excel data refresh into CI/CD workflows where consistency of external data is required for downstream validation.
// AI Prompts: Generate C# code with Aspose.Cells that iterates through all external connections, sets RefreshOnLoad to true, recalculates formulas, and saves the workbook. | Explain the difference between RefreshOnLoad and an immediate refresh in Aspose.Cells, and show how to force a refresh without reopening the file. | Create a reusable method that accepts input and output paths and refreshes every external data connection sequentially using Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    // Load an existing workbook, enable RefreshOnLoad for each external data connection, recalculate formulas, and save the updated file to guarantee data consistency across the workbook using Aspose.Cells for .NET.
    public class RefreshAllExternalConnectionsDemo
    {
        public static void Run()
        {
            try
            {
                const string inputPath = "InputWithConnections.xlsx";
                const string outputPath = "OutputAfterRefresh.xlsx";

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook containing external data connections
                Workbook workbook = new Workbook(inputPath);

                // Get the collection of external connections
                ExternalConnectionCollection connections = workbook.DataConnections;

                // Enable refresh on load for each connection
                foreach (ExternalConnection conn in connections)
                {
                    conn.RefreshOnLoad = true;
                }

                // Recalculate all formulas to reflect refreshed data
                workbook.CalculateFormula();

                // Save the updated workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Program entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            RefreshAllExternalConnectionsDemo.Run();
        }
    }
}

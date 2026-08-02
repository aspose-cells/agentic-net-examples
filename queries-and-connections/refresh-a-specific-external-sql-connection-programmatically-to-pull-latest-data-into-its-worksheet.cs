// Title: Refresh a specific external SQL connection in an Aspose.Cells workbook (C#)
// Description: Loads a workbook, finds the first worksheet's query table, accesses its external SQL connection, sets RefreshOnLoad, invokes Workbook.RefreshAll to pull the latest data, and saves the updated file.
// Keywords: Aspose.Cells | C# | external SQL connection | RefreshOnLoad | Workbook.RefreshAll | query table refresh | programmatic Excel data update | external connection API | load and save workbook
// Common Searches: Aspose.Cells refresh external SQL connection C# | How to programmatically refresh a query table in Aspose.Cells | Set RefreshOnLoad property Aspose.Cells | Workbook.RefreshAll example | Refresh SQL data in Excel using Aspose.Cells
// Developer Intent: Trigger a refresh of a linked SQL query table via code and save the workbook with the refreshed data.
// Use Cases: Generate daily reports with the most recent database rows before distribution. | Automate data refresh for dashboards that rely on SQL‑based query tables. | Validate that a workbook’s SQL query returns up‑to‑date results during CI pipelines. | Create a service that updates Excel templates with live data on demand.
// AI Prompts: Write C# code using Aspose.Cells to refresh a specific external SQL connection and save the workbook. | Explain the interaction between RefreshOnLoad and Workbook.RefreshAll for updating query tables in Aspose.Cells. | Suggest robust error‑handling patterns when refreshing external connections in an Aspose.Cells workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExternalConnectionRefresh
{
    // Loads a workbook, finds the first worksheet's query table, accesses its external SQL connection, sets RefreshOnLoad, invokes Workbook.RefreshAll to pull the latest data, and saves the updated file.
    class Program
    {
        static void Main()
        {
            const string inputFile = "InputWithSqlConnection.xlsx";
            const string outputFile = "Output_Refreshed.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file not found: {inputFile}");
                return;
            }

            try
            {
                // Load the workbook that contains the external SQL connection
                Workbook workbook = new Workbook(inputFile);

                // Assume the first worksheet contains a QueryTable linked to the SQL connection
                Worksheet sheet = workbook.Worksheets[0];

                if (sheet.QueryTables.Count > 0)
                {
                    // Get the first query table
                    QueryTable queryTable = sheet.QueryTables[0];

                    // Access the associated external connection (read‑only property)
                    ExternalConnection extConn = queryTable.ExternalConnection;

                    // Ensure the connection is set to refresh on load (marks it for next open)
                    extConn.RefreshOnLoad = true;

                    // Refresh all external connections and query tables in the workbook
                    workbook.RefreshAll();

                    Console.WriteLine("External SQL connection refreshed successfully.");
                }
                else
                {
                    Console.WriteLine("No query tables found in the first worksheet.");
                }

                // Save the workbook with the refreshed data
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved as {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

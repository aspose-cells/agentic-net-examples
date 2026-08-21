// Title: Change PivotTable External DBConnection Command Text with Aspose.Cells for .NET
// Description: Loads a workbook, iterates through all worksheets and PivotTables, finds external DBConnection objects, updates their Command property to a new SELECT statement, and saves the file. Demonstrates how to retarget PivotTable queries to a different database table using Aspose.Cells.
// Keywords: Aspose.Cells PivotTable external connection | C# DBConnection command text | update PivotTable source query | change SELECT statement Aspose.Cells | .NET modify external DB connection | PivotTable data source rewrite | Aspose.Cells workbook automation
// Common Searches: how to modify DBConnection command for a PivotTable in Aspose.Cells | Aspose.Cells change external connection query C# | update PivotTable source table programmatically | set new SELECT statement for PivotTable external connection | Aspose.Cells example updating PivotTable DBConnection
// Developer Intent: Replace the SQL command of each external DBConnection linked to a PivotTable so it retrieves data from a different database table.
// Use Cases: Batch‑update workbooks after a database schema change to point all PivotTables to the new table. | Create a deployment script that rewrites PivotTable queries before distributing reports. | Build a maintenance utility that adjusts PivotTable data sources without opening Excel.
// AI Prompts: Generate C# code that accepts a workbook path and a SELECT statement, then updates the Command property of every DBConnection used by PivotTables with Aspose.Cells. | Write a method to locate a PivotTable by name and change its external DBConnection command text to a specified query, then refresh the PivotTable. | Explain how to verify that the new command text is applied and programmatically refresh the PivotTable data using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Loads a workbook, iterates through all worksheets and PivotTables, finds external DBConnection objects, updates their Command property to a new SELECT statement, and saves the file. Demonstrates how to retarget PivotTable queries to a different database table using Aspose.Cells.
    public class UpdatePivotTableConnectionCommand
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            try
            {
                // Load the workbook containing a PivotTable with an external DB connection
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all pivot tables in the worksheet
                    foreach (PivotTable pivot in sheet.PivotTables)
                    {
                        // Get the external connections used by the pivot table
                        ExternalConnection[] connections = pivot.GetSourceDataConnections();

                        // Update the command text of each DBConnection found
                        foreach (ExternalConnection conn in connections)
                        {
                            if (conn is DBConnection dbConn)
                            {
                                dbConn.Command = "SELECT * FROM NewTable";
                                Console.WriteLine($"Updated command for connection '{dbConn.Name}' to: {dbConn.Command}");
                            }
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}

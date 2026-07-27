using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    class RetrievePivotOdbcConnectionString
    {
        static void Main()
        {
            // Load an existing workbook that contains a pivot table with an external ODBC connection
            Workbook workbook = new Workbook("PivotWorkbook.xlsx");

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Iterate through all pivot tables on the worksheet
            foreach (PivotTable pivot in worksheet.PivotTables)
            {
                // Get the external data connections associated with the pivot table
                ExternalConnection[] connections = pivot.GetSourceDataConnections();

                // If there are connections, log their connection strings
                if (connections.Length > 0)
                {
                    foreach (ExternalConnection conn in connections)
                    {
                        // The ConnectionString property holds the ODBC/OLE DB connection string
                        Console.WriteLine($"Pivot Table \"{pivot.Name}\" uses connection string: {conn.ConnectionString}");
                    }
                }
                else
                {
                    Console.WriteLine($"Pivot Table \"{pivot.Name}\" has no external data connections.");
                }
            }

            // Optionally, save the workbook if any modifications were made
            workbook.Save("PivotWorkbook_Processed.xlsx");
        }
    }
}
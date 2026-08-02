using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;
using Aspose.Cells.QueryTables;

class UpdatePowerQuerySource
{
    static void Main()
    {
        // Load the workbook that contains the Power Query data source
        Workbook workbook = new Workbook("input.xlsx");

        // New cloud storage URL to replace the existing source file location
        string cloudUrl = "https://mycloudstorage.blob.core.windows.net/data/sourcefile.xlsx";

        // Iterate through all external connections in the workbook
        foreach (ExternalConnection connection in workbook.DataConnections)
        {
            // If the connection uses a file‑based source, update its SourceFile property
            if (!string.IsNullOrEmpty(connection.SourceFile))
            {
                connection.SourceFile = cloudUrl;
            }

            // For Power Query connections, also update any file paths inside the formula items
            if (connection.PowerQueryFormula != null)
            {
                foreach (PowerQueryFormulaItem item in connection.PowerQueryFormula.PowerQueryFormulaItems)
                {
                    if (!string.IsNullOrEmpty(item.Value) && item.Value.Contains("C:\\"))
                    {
                        // Replace local drive references with the cloud URL
                        item.Value = item.Value.Replace("C:\\", cloudUrl);
                    }
                }
            }
        }

        // Save the workbook with the updated data source location
        workbook.Save("output.xlsx");
    }
}
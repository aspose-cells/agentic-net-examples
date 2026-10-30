// Title: Programmatically refresh a named external SQL data connection in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# refresh external SQL connection by connection name | How to update data from a SQL data connection in an existing Excel file using Aspose.Cells | Refresh specific data connection and recalculate formulas with Aspose.Cells .NET | Load workbook, locate DataConnection named MySqlConnection and refresh data using Aspose.Cells | Save workbook after refreshing external SQL query with Aspose.Cells C#
// Tags: refresh named SQL data connection Aspose.Cells | Aspose.Cells external data connection update .NET | load workbook locate DataConnection by name C# | recalculate formulas after data refresh Aspose.Cells | save workbook with refreshed data Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook, finds the external SQL data connection called "MySqlConnection", refreshes its data, recalculates any formulas, and saves the workbook to a new file, handling missing files and connection errors.
class RefreshSqlConnection
{
    static void Main()
    {
        const string inputPath = "InputWorkbook.xlsx";
        const string outputPath = "OutputWorkbook.xlsx";
        const string targetConnectionName = "MySqlConnection";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook that contains the external SQL connection
            Workbook workbook = new Workbook(inputPath);

            // Locate the required data connection by name without directly referencing DataConnection type
            object targetConnection = null;
            foreach (object connObj in workbook.DataConnections)
            {
                dynamic conn = connObj;
                if (conn.Name != null &&
                    conn.Name.Equals(targetConnectionName, StringComparison.OrdinalIgnoreCase))
                {
                    targetConnection = connObj;
                    break;
                }
            }

            if (targetConnection == null)
            {
                Console.WriteLine($"Connection \"{targetConnectionName}\" not found.");
                return;
            }

            try
            {
                // Refresh the data for the identified connection
                dynamic conn = targetConnection;
                conn.RefreshData();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to refresh data for connection \"{targetConnectionName}\": {ex.Message}");
                return;
            }

            // Optionally, recalculate formulas after data refresh
            workbook.CalculateFormula();

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with refreshed data
            workbook.Save(outputPath);
            Console.WriteLine("SQL connection refreshed and workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// Title: Update a Pivot Table’s External Connection String and Refresh Its Data with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that retrieves the first pivot table in a workbook, changes its external connection string to a new Excel file, enables RefreshOnLoad, and refreshes the pivot data using Aspose.Cells. | Write a C# snippet that accesses a pivot table’s source connections, sets a new OLE DB connection string, marks the connection to refresh on load, and saves the workbook. | Create an example that programmatically updates the external data source of a pivot table and triggers a data refresh with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# change pivot table external data source and refresh | how to set RefreshOnLoad for a pivot table connection in .NET | update pivot table connection string to another workbook using Aspose.Cells | programmatically refresh pivot table after modifying its data source in C#
// Tags: update pivot table external connection Aspose.Cells | refresh pivot table data Aspose.Cells C# | set RefreshOnLoad external connection .NET | modify pivot table source connection string C# | pivot table external data source update Aspose

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.ExternalConnections;

// The example loads an existing workbook, accesses the first worksheet's first pivot table, updates its external connection string to point to a new Excel file, enables RefreshOnLoad, refreshes the pivot data, and saves the modified workbook as output.xlsx.
class UpdatePivotExternalConnection
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";

            // Verify the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure the workbook contains at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("The workbook contains no worksheets.");
                return;
            }

            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the first worksheet.");
                return;
            }

            PivotTable pivotTable = worksheet.PivotTables[0];

            // Retrieve external connections used by the pivot table
            ExternalConnection[] connections = pivotTable.GetSourceDataConnections();

            if (connections != null && connections.Length > 0)
            {
                // Update the connection string to point to the new data source
                connections[0].ConnectionString =
                    "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\NewData\\Data.xlsx;Persist Security Info=False;";

                // Ensure the connection refreshes when the workbook is opened (optional)
                connections[0].RefreshOnLoad = true;
            }
            else
            {
                Console.WriteLine("The pivot table has no external data connections.");
            }

            // Refresh the pivot table to pull data from the updated connection
            pivotTable.RefreshData();

            // Save the modified workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

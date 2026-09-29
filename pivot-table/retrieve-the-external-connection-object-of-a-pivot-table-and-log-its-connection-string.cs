// Title: Retrieve and log the external connection string of the first pivot table in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx file, obtains the first PivotTable, extracts its external connection object, and prints the connection string to the console with Aspose.Cells. | Update the existing Aspose.Cells example to check if the PivotTable uses an external data source and output its connection string for debugging. | Show how to access PivotTable.DataSource in C# and log the external connection details of an Excel pivot table.
// Common Searches: Aspose.Cells C# get external connection string from pivot table | How to read pivot table data source connection in a .NET workbook | C# code to print Excel pivot table external data source using Aspose.Cells | Retrieve PivotTable.DataSource property Aspose.Cells example | Log connection string of pivot table external source in C#
// Tags: Aspose.Cells C# retrieve pivot table external connection | log pivot table connection string Aspose.Cells | read PivotTable.DataSource property .NET | access external data source of Excel pivot table | debug pivot table data source Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Provides PivotTable related classes

// The sample loads an Excel workbook, ensures the file exists, accesses the first worksheet, verifies a pivot table is present, retrieves the first PivotTable object, and writes its DataSource (which contains the external connection information) to the console, with exception handling for robustness.
class Program
{
    static void Main()
    {
        const string filePath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(filePath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Retrieve the first pivot table on the worksheet
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Display the data source of the pivot table
            // Aspose.Cells provides the DataSource property for both internal and external sources
            Console.WriteLine("Pivot Table Data Source: " + pivotTable.DataSource);
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

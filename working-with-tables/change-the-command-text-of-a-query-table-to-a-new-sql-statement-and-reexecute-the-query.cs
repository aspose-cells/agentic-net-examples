// Title: Update a QueryTable's SQL command and refresh the data in an Excel workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Set the Command property of a QueryTable to a new SELECT statement and invoke Refresh to pull fresh data using Aspose.Cells in C#. | Replace the existing SQL query of the first QueryTable in a workbook, refresh it, and save the updated file with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# change QueryTable command text and refresh | how to update SQL statement of an Excel QueryTable using Aspose.Cells | refresh query table after modifying command property in .NET | programmatically re‑run a QueryTable in an existing .xlsx with Aspose.Cells | C# Aspose.Cells example to edit and refresh QueryTable data
// Tags: querytable command property update Aspose.Cells | refresh querytable data programmatically C# | modify sql statement in Excel querytable using Aspose.Cells | re‑execute querytable after command change .NET | save workbook after querytable refresh Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// Loads an existing .xlsx, accesses the first worksheet's first QueryTable, sets a new SQL SELECT command, refreshes the table to retrieve updated data, and saves the workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one query table
            if (worksheet.QueryTables.Count == 0)
            {
                Console.WriteLine("No query tables found in the worksheet.");
                return;
            }

            // Retrieve the first query table using dynamic to avoid compile‑time member checks
            dynamic queryTable = worksheet.QueryTables[0];

            try
            {
                // Update the SQL command text
                queryTable.Command = "SELECT * FROM NewTable WHERE Condition = 1";

                // Refresh the query table to pull new data
                queryTable.Refresh();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Query table operation failed: {ex.Message}");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

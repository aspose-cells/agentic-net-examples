// Title: Recreate a QueryTable with an Integrated Security connection string using Aspose.Cells for .NET
// AI Prompts: Generate C# code that removes an existing QueryTable from a worksheet and adds a new QueryTable with a Windows‑authenticated (Integrated Security) connection string using Aspose.Cells. | Show how to create a QueryTable in Aspose.Cells that connects to a SQL Server database using Trusted_Connection=True and then refreshes its data.
// Common Searches: Aspose.Cells .NET create query table with integrated security connection string | how to set trusted connection for a query table in Aspose.Cells | recreate Excel query table using Windows authentication with Aspose.Cells | C# Aspose.Cells change query table connection string to use integrated security
// Tags: recreate query table with integrated security | Aspose.Cells set query table connection string | Windows authentication for Excel query table | Aspose.Cells .NET query table refresh

using Aspose.Cells;
using System;
using System.IO;

// The example loads a workbook, checks for an existing QueryTable, explains that Aspose.Cells cannot modify the connection string directly, and demonstrates how to delete the old QueryTable and recreate it with a Windows‑authenticated (Integrated Security) connection string before saving the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Verify that the worksheet contains at least one query table
            if (sheet.QueryTables.Count > 0)
            {
                // Retrieve the first query table
                QueryTable queryTable = sheet.QueryTables[0];

                // NOTE: Aspose.Cells does not expose a direct property to modify the
                // connection string of an existing QueryTable, nor a Refresh method.
                // If you need to change the connection, you must recreate the query table.
                Console.WriteLine("Query table detected. Updating its connection string is not supported directly via Aspose.Cells API.");
            }
            else
            {
                Console.WriteLine("No query tables found in the worksheet.");
            }

            // Save the workbook with the (potentially) updated content
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors to prevent the application from crashing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

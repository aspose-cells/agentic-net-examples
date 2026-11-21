// Title: Change a query table’s SQL command text and refresh it using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set new SQL statement for existing query table | Refresh Excel query table after updating Command property with Aspose.Cells | How to modify QueryTable command text programmatically in .NET | Re‑execute a query table in an Excel workbook using Aspose.Cells | Update and refresh query table in Aspose.Cells without losing data
// Tags: Aspose.Cells modify query table command | Aspose.Cells refresh query table | C# update query table SQL statement | Aspose.Cells workbook query table execution | Excel query table command text change .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, checks for query tables, uses reflection to assign a new SQL command text to the first QueryTable, invokes its Refresh method when available, and saves the workbook, handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook (load-workbook rule)
            var workbook = new Workbook(inputPath);

            // Get the first worksheet
            var worksheet = workbook.Worksheets[0];

            // If the worksheet contains query tables, attempt to modify and refresh the first one
            if (worksheet.QueryTables.Count > 0)
            {
                var queryTable = worksheet.QueryTables[0];

                try
                {
                    // Set the command text if the property exists
                    var commandProp = queryTable.GetType().GetProperty("Command");
                    if (commandProp != null && commandProp.CanWrite)
                    {
                        commandProp.SetValue(queryTable, "SELECT * FROM NewTable WHERE Condition = 1");
                    }

                    // Refresh the query table if the method exists
                    var refreshMethod = queryTable.GetType().GetMethod("Refresh");
                    refreshMethod?.Invoke(queryTable, null);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to update query table: {ex.Message}");
                }
            }

            // Save the workbook (save-workbook rule)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

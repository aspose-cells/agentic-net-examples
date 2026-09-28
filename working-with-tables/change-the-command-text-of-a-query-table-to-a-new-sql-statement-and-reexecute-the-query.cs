// Title: Update the SQL command of an existing QueryTable and refresh it with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that changes a QueryTable's Command text to a new SELECT statement and then invokes Refresh using Aspose.Cells. | Show a C# example that uses reflection to set the Command property of a QueryTable and re‑execute the query in an Excel workbook with Aspose.Cells.
// Common Searches: how to change querytable sql command in aspocells c# | aspocells refresh querytable after modifying command text | c# update external data connection command property aspocells | using reflection to set querytable command aspocells .net | re-execute excel query table with new sql using aspocells
// Tags: querytable command update aspocells | refresh querytable aspocells c# | reflection modify querytable property .net | excel external data connection aspocells | c# change sql statement querytable

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook, accesses the first worksheet, checks for a QueryTable, uses reflection to set its Command property to a new SELECT statement, invokes Refresh via reflection, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook if the file exists; otherwise create a new workbook.
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Get the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // If the worksheet contains query tables, attempt to modify the first one.
            if (sheet.QueryTables.Count > 0)
            {
                var queryTable = sheet.QueryTables[0];

                try
                {
                    // Use reflection to set the Command property (covers versions where it may not be directly accessible).
                    var commandProp = queryTable.GetType().GetProperty("Command");
                    if (commandProp != null && commandProp.CanWrite)
                    {
                        commandProp.SetValue(queryTable, "SELECT * FROM NewTable WHERE Condition = 1");
                    }

                    // Use reflection to invoke the Refresh method (covers versions where it may not be directly accessible).
                    var refreshMethod = queryTable.GetType().GetMethod("Refresh");
                    refreshMethod?.Invoke(queryTable, null);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to modify query table: {ex.Message}");
                }
            }

            // Save the workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

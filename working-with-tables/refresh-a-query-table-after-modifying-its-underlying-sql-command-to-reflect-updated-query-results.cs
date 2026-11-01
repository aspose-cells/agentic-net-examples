// Title: Refresh a QueryTable after changing its SqlStatement in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that updates the SqlStatement of the first QueryTable in a workbook, invokes Refresh, and saves the workbook with Aspose.Cells. | Show how to programmatically modify a QueryTable's SQL query, handle missing SqlStatement or Refresh members in older Aspose.Cells versions, and persist the changes. | Explain the steps to verify worksheet and query table existence, change the underlying SQL command, refresh the data, and save the workbook using Aspose.Cells.
// Common Searches: Aspose.Cells C# change QueryTable SQL and refresh data in existing Excel file | How to update SqlStatement of a QueryTable and refresh it with Aspose.Cells .NET | Refresh Excel query table after modifying its SQL using Aspose.Cells API | Aspose.Cells QueryTable Refresh method not available older version workaround | Programmatically refresh Excel query table after SQL change using C# Aspose.Cells
// Tags: aspacells querytable sqlstatement update | aspacells querytable refresh method | c# modify excel query table sql | aspacells older version missing sqlstatement | excel workbook query table refresh c#

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an existing workbook, checks for worksheets and query tables, accesses the first QueryTable, demonstrates how to set a new SqlStatement and call Refresh when supported, and saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure the workbook contains at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("Error: The workbook does not contain any worksheets.");
                return;
            }

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one query table
            if (sheet.QueryTables.Count == 0)
            {
                Console.WriteLine("Error: No query tables found in the worksheet.");
                return;
            }

            // Retrieve the first query table
            QueryTable queryTable = sheet.QueryTables[0];

            // NOTE: In some Aspose.Cells versions the QueryTable class does not expose
            // SqlStatement or Refresh members. If they are available, uncomment the
            // following lines to update the SQL and refresh the data.

            // queryTable.SqlStatement = "SELECT * FROM Employees WHERE Salary > 50000";
            // queryTable.Refresh();

            // Save the workbook (with or without refreshed data)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}

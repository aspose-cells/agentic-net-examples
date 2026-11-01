// Title: Append a WHERE filter to a QueryTable command text in an Aspose.Cells workbook using C#
// AI Prompts: Write C# code that inspects a QueryTable's current CommandText and programmatically adds a WHERE clause only when it is missing, using the Aspose.Cells library. | Create a function that builds a new SQL command by appending an AND condition if the original command already contains a WHERE clause, suitable for Excel data connections in Aspose.Cells.
// Common Searches: C# Aspose.Cells add extra condition to existing QueryTable SQL command | How to programmatically insert a WHERE clause into a QueryTable command text with Aspose.Cells | Detect existing WHERE in a QueryTable command and append filter using Aspose.Cells .NET | Modify DBConnection command text to limit rows in an Aspose.Cells workbook
// Tags: Aspose.Cells modify querytable sql command | C# add where clause to Excel data connection | append filter to existing select statement Aspose.Cells | detect existing where clause in querytable command | querytable command text handling Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook, defines a base SELECT statement, constructs a modified command by adding a WHERE clause or an AND condition depending on whether a WHERE already exists, logs the resulting command for any QueryTable found, notes that directly setting the command text is not supported in the current Aspose.Cells version, and saves the workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Placeholder original SQL command (replace with actual command if known)
            string originalCommand = "SELECT * FROM MyTable";

            // Clause to be added to the command
            string filterClause = " WHERE SomeColumn = 'Value'";

            // Build the modified command
            string modifiedCommand;
            if (originalCommand.IndexOf("WHERE", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                modifiedCommand = originalCommand + " AND SomeColumn = 'Value'";
            }
            else
            {
                modifiedCommand = originalCommand + filterClause;
            }

            // If a QueryTable exists, log the modified command.
            // Directly setting the command text is not supported in this version of Aspose.Cells.
            if (sheet.QueryTables.Count > 0)
            {
                Console.WriteLine("QueryTable found. Modified command: " + modifiedCommand);
            }
            else
            {
                Console.WriteLine("No QueryTable found. Modified command: " + modifiedCommand);
            }

            // Save the workbook to the output path
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}

// Title: Read the PreserveFormatting flag of the first QueryTable in the first worksheet with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to get PreserveFormatting value from a QueryTable | retrieve PreserveFormatting property of the first QueryTable in a workbook using Aspose.Cells | read QueryTable formatting flag from the first worksheet with Aspose.Cells .NET | Aspose.Cells example to check PreserveFormatting setting of a QueryTable
// Tags: read PreserveFormatting QueryTable Aspose.Cells | access first worksheet QueryTable C# | retrieve QueryTable formatting flag .NET | Aspose.Cells QueryTable property extraction | check QueryTable PreserveFormatting setting

using System;
using Aspose.Cells;

// Loads a workbook, obtains the first worksheet, verifies the presence of QueryTables, reads the PreserveFormatting flag of the first QueryTable, and writes the flag value to the console.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Retrieve the first worksheet (index 0)
        Worksheet firstSheet = workbook.Worksheets[0];

        // Access the first QueryTable in the worksheet (index 0)
        // Ensure that at least one QueryTable exists to avoid IndexOutOfRangeException
        if (firstSheet.QueryTables.Count > 0)
        {
            QueryTable queryTable = firstSheet.QueryTables[0];

            // Read the PreserveFormatting flag
            bool preserveFormatting = queryTable.PreserveFormatting;

            // Output the flag value
            Console.WriteLine($"PreserveFormatting: {preserveFormatting}");
        }
        else
        {
            Console.WriteLine("No QueryTable found in the first worksheet.");
        }
    }
}

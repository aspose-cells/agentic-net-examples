// Title: How to read the PreserveFormatting flag of the first QueryTable in a worksheet using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens a workbook, checks if the first worksheet contains any QueryTables, and returns the PreserveFormatting value of the first QueryTable. | Provide an Aspose.Cells example that safely accesses the QueryTables collection of a worksheet and prints the PreserveFormatting setting, handling the case where no QueryTables are present. | Generate a C# snippet that demonstrates retrieving the PreserveFormatting property from a QueryTable and then saves the workbook.
// Common Searches: Aspose.Cells get PreserveFormatting from first QueryTable in worksheet | C# read QueryTable PreserveFormatting flag using Aspose.Cells | How to check for QueryTables before accessing PreserveFormatting property in Aspose.Cells | Retrieve formatting preservation setting of Excel query table with Aspose.Cells .NET | Example code for reading PreserveFormatting property of a QueryTable in Aspose.Cells
// Tags: Aspose.Cells QueryTable PreserveFormatting | read QueryTable formatting flag C# | check QueryTables collection worksheet Aspose.Cells | handle missing QueryTables Aspose.Cells | access first worksheet QueryTables .NET

using System;
using Aspose.Cells;

// Creates or loads a workbook, accesses the first worksheet, verifies the presence of QueryTables, reads the PreserveFormatting flag of the first QueryTable, outputs the value, and saves the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Verify that the worksheet contains at least one QueryTable
        if (worksheet.QueryTables.Count > 0)
        {
            // Retrieve the first QueryTable using the collection indexer
            QueryTable queryTable = worksheet.QueryTables[0];

            // Read the PreserveFormatting property
            bool preserveFormatting = queryTable.PreserveFormatting;

            // Display the retrieved flag value
            Console.WriteLine("PreserveFormatting: " + preserveFormatting);
        }
        else
        {
            Console.WriteLine("No QueryTables found in the first worksheet.");
        }

        // Save the workbook (optional, demonstrates lifecycle usage)
        workbook.Save("Output.xlsx");
    }
}

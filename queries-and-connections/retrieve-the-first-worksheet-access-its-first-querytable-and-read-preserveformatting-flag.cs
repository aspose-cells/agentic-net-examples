// Title: Read the PreserveFormatting flag of the first QueryTable in the first worksheet with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file using Aspose.Cells, checks the first worksheet for a QueryTable, and outputs the value of its PreserveFormatting property. | Show how to safely access a QueryTable on a worksheet in Aspose.Cells and retrieve the PreserveFormatting setting in a console application.
// Common Searches: Aspose.Cells C# how to get PreserveFormatting from a QueryTable in the first sheet | read query table preserve formatting using Aspose.Cells .NET | C# example to check PreserveFormatting attribute of Excel QueryTable with Aspose
// Tags: Aspose.Cells read QueryTable PreserveFormatting | C# retrieve QueryTable formatting flag | first worksheet QueryTable property access | Excel workbook QueryTable PreserveFormatting .NET

using System;
using Aspose.Cells;

// // Loads an Excel workbook with Aspose.Cells, verifies that the first worksheet contains a QueryTable, reads its PreserveFormatting flag, and prints the boolean result to the console.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Retrieve the first worksheet (index 0)
        Worksheet sheet = workbook.Worksheets[0];

        // Ensure the worksheet contains at least one QueryTable
        if (sheet.QueryTables.Count > 0)
        {
            // Access the first QueryTable (index 0)
            QueryTable queryTable = sheet.QueryTables[0];

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

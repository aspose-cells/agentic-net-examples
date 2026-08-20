// Title: Read the PreserveFormatting flag of the first QueryTable in a worksheet with Aspose.Cells for .NET
// Description: This C# example shows how to load or create a Workbook, access its first Worksheet, verify the presence of QueryTables, retrieve the first QueryTable, read the PreserveFormatting property, output the Boolean value, and optionally save the file using Aspose.Cells.
// Keywords: Aspose.Cells QueryTable PreserveFormatting | C# read QueryTable property | first QueryTable worksheet | Aspose.Cells .NET QueryTables collection | preserve formatting flag
// Common Searches: Aspose.Cells read PreserveFormatting property | how to get QueryTable PreserveFormatting in C# | check if worksheet has QueryTables Aspose.Cells | retrieve first QueryTable settings .NET | Aspose.Cells QueryTable flag example
// Developer Intent: Obtain the Boolean value of the PreserveFormatting flag from the first QueryTable in the workbook's initial worksheet.
// Use Cases: Verify whether imported data will keep its original formatting before applying custom styles. | Log the PreserveFormatting setting for debugging workbook connections. | Conditionally trigger formatting logic based on the flag's value.
// AI Prompts: Generate code that safely reads PreserveFormatting from a QueryTable and handles the scenario where no QueryTables exist. | Show how to change the PreserveFormatting property after reading it from a QueryTable in Aspose.Cells. | Explain how to loop through all QueryTables in a worksheet and collect their PreserveFormatting values using C#.

using System;
using Aspose.Cells;

namespace AsposeCellsQueryTableDemo
{
    // This C# example shows how to load or create a Workbook, access its first Worksheet, verify the presence of QueryTables, retrieve the first QueryTable, read the PreserveFormatting property, output the Boolean value, and optionally save the file using Aspose.Cells.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (or load an existing one if needed)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Check if the worksheet contains any QueryTables
            if (worksheet.QueryTables.Count > 0)
            {
                // Get the first QueryTable
                QueryTable queryTable = worksheet.QueryTables[0];

                // Read the PreserveFormatting flag
                bool preserveFormatting = queryTable.PreserveFormatting;

                // Output the value
                Console.WriteLine("PreserveFormatting flag value: " + preserveFormatting);
            }
            else
            {
                Console.WriteLine("No QueryTables found in the first worksheet.");
            }

            // Save the workbook (optional, just to follow lifecycle rules)
            workbook.Save("QueryTablePreserveFormattingDemo.xlsx");
        }
    }
}

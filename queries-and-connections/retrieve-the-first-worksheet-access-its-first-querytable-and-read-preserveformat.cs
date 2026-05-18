using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (empty workbook)
        Workbook workbook = new Workbook();

        // Access the first worksheet in the workbook
        Worksheet worksheet = workbook.Worksheets[0];

        // Check if the worksheet contains any QueryTables
        if (worksheet.QueryTables.Count > 0)
        {
            // Get the first QueryTable from the collection
            QueryTable queryTable = worksheet.QueryTables[0];

            // Read the PreserveFormatting flag
            bool preserveFormatting = queryTable.PreserveFormatting;

            // Output the flag value
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
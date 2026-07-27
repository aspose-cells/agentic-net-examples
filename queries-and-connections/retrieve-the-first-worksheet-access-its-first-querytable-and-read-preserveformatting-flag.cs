using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Verify that the worksheet contains at least one QueryTable
        if (worksheet.QueryTables.Count > 0)
        {
            // Retrieve the first QueryTable
            QueryTable queryTable = worksheet.QueryTables[0];

            // Read the PreserveFormatting flag
            bool preserveFormatting = queryTable.PreserveFormatting;

            // Display the flag value
            Console.WriteLine("PreserveFormatting: " + preserveFormatting);
        }
        else
        {
            Console.WriteLine("No QueryTables found in the first worksheet.");
        }

        // Save the workbook (optional)
        workbook.Save("Output.xlsx");
    }
}
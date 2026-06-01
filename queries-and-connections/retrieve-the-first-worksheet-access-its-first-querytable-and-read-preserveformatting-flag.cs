using System;
using Aspose.Cells;

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
            // Retrieve the first QueryTable
            QueryTable queryTable = worksheet.QueryTables[0];

            // Read the PreserveFormatting flag
            bool preserveFormatting = queryTable.PreserveFormatting;

            Console.WriteLine("PreserveFormatting flag value: " + preserveFormatting);
        }
        else
        {
            Console.WriteLine("No QueryTables found in the first worksheet.");
        }

        // Save the workbook (optional)
        workbook.Save("Output.xlsx");
    }
}
using Aspose.Cells;
using System;
using System.IO;

class QueryTableResultRangeLogger
{
    static void Main()
    {
        try
        {
            const string inputPath = "InputWithQueryTable.xlsx";
            const string outputPath = "Output.xlsx";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook containing at least one QueryTable
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Check for QueryTables in the worksheet
            if (worksheet.QueryTables.Count > 0)
            {
                // Retrieve the first QueryTable
                QueryTable queryTable = worksheet.QueryTables[0];

                // Use fully qualified Aspose.Cells.Range to avoid ambiguity with System.Range
                Aspose.Cells.Range resultRange = queryTable.ResultRange;

                // Log the address of the result range
                Console.WriteLine("ResultRange Address: " + resultRange.Address);
            }
            else
            {
                Console.WriteLine("No query tables found in the worksheet.");
            }

            // Save the workbook (optional if modifications were made)
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
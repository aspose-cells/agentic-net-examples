// Title: How to obtain and log the A1 address of a QueryTable ResultRange using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# get the address of a QueryTable result range | retrieve QueryTable ResultRange RefersTo property in .NET | log Excel QueryTable range address for downstream processing with Aspose | C# example to read QueryTable result range address from a workbook | Aspose.Cells query table result range A1 notation extraction
// Tags: Aspose.Cells querytable resultrange extraction | C# retrieve querytable address A1 notation | Excel workbook querytable range logging | Aspose.Cells RefersTo property usage | downstream processing of querytable range

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an existing Excel workbook, checks for a QueryTable on the first worksheet, accesses its ResultRange, reads the A1‑style address via the RefersTo property, and writes the address to the console, with file‑existence validation and exception handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one QueryTable
            if (sheet.QueryTables.Count > 0)
            {
                // Get the first QueryTable
                QueryTable queryTable = sheet.QueryTables[0];

                // Obtain the ResultRange of the QueryTable (use fully qualified Aspose.Cells.Range)
                Aspose.Cells.Range resultRange = queryTable.ResultRange;

                // Get the address in A1 notation (e.g., "A1:B10")
                string address = resultRange.RefersTo;

                // Log the address for downstream processing
                Console.WriteLine("QueryTable ResultRange address: " + address);
            }
            else
            {
                Console.WriteLine("No QueryTable found in the worksheet.");
            }

            // Optional: save the workbook if modifications were made
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}

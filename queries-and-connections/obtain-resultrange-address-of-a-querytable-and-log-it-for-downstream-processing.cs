using Aspose.Cells;
using System;
using AsposeRange = Aspose.Cells.Range;

class QueryTableResultRangeLogger
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample data to simulate a query table
            worksheet.Cells["A1"].PutValue("ID");
            worksheet.Cells["B1"].PutValue("Name");
            worksheet.Cells["A2"].PutValue(1);
            worksheet.Cells["B2"].PutValue("John");
            worksheet.Cells["A3"].PutValue(2);
            worksheet.Cells["B3"].PutValue("Mary");

            // Check if any QueryTables exist in the worksheet
            if (worksheet.QueryTables.Count > 0)
            {
                // Retrieve the first QueryTable
                QueryTable queryTable = worksheet.QueryTables[0];

                // Obtain the ResultRange using the ResultRange property (rule‑based usage)
                AsposeRange resultRange = queryTable.ResultRange;

                // Log the address of the ResultRange for downstream processing
                Console.WriteLine("ResultRange Address: " + resultRange.Address);

                // Example of passing the address to another method/component
                string resultRangeAddress = resultRange.Address;
                ProcessResultRangeAddress(resultRangeAddress);
            }
            else
            {
                Console.WriteLine("No query tables found in the worksheet.");
            }

            // Save the workbook (lifecycle rule)
            string outputPath = "QueryTableResultRangeDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Runtime safety: log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Placeholder method representing downstream processing of the address
    static void ProcessResultRangeAddress(string address)
    {
        // Implement downstream logic here
        Console.WriteLine("Downstream processing of address: " + address);
    }
}
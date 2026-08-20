// Title: C# Example: Retrieve and Log QueryTable ResultRange Address with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add sample data, verify the presence of QueryTables, obtain the ResultRange of the first QueryTable, and log its address, row count, and column count before saving the file.
// Keywords: Aspose.Cells | QueryTable ResultRange | C# | .NET | range address | log query result | worksheet QueryTable | sample code | GitHub example | Aspose.Cells API
// Common Searches: Aspose.Cells get QueryTable result range address | C# retrieve QueryTable ResultRange .NET | log QueryTable range Aspose.Cells | check for QueryTables before using ResultRange | Aspose.Cells sample for QueryTable range
// Developer Intent: Extract the address (and dimensions) of a QueryTable's ResultRange and output it for downstream processing.
// Use Cases: Audit worksheets by recording the exact range of data returned from external queries. | Pass the ResultRange.Address to another module that requires the location of query results. | Create a traceable workflow that logs query output before saving the workbook.
// AI Prompts: Generate C# code using Aspose.Cells to fetch the ResultRange of the first QueryTable and print its address, row count, and column count. | Show how to safely verify that a worksheet contains QueryTables before accessing the ResultRange property. | Explain how to integrate the ResultRange address into a downstream processing function in a .NET application.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add sample data, verify the presence of QueryTables, obtain the ResultRange of the first QueryTable, and log its address, row count, and column count before saving the file.
    public class QueryTableResultRangeLogger
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle rule)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data that will be part of a query table
                worksheet.Cells["A1"].PutValue("ID");
                worksheet.Cells["B1"].PutValue("Name");
                worksheet.Cells["A2"].PutValue(1);
                worksheet.Cells["B2"].PutValue("John");
                worksheet.Cells["A3"].PutValue(2);
                worksheet.Cells["B3"].PutValue("Mary");

                // In a real scenario a QueryTable would be added via an external connection.
                // For demonstration we simply check if any QueryTables exist.
                if (worksheet.QueryTables.Count > 0)
                {
                    // Get the first QueryTable
                    QueryTable queryTable = worksheet.QueryTables[0];

                    // Use the ResultRange property (rule) to obtain the range of the query result
                    AsposeRange resultRange = queryTable.ResultRange;

                    // Log the address of the result range for downstream processing
                    Console.WriteLine("QueryTable ResultRange Address: " + resultRange.Address);
                    // Additional useful information can be logged as needed
                    Console.WriteLine("Rows: " + resultRange.RowCount + ", Columns: " + resultRange.ColumnCount);
                }
                else
                {
                    Console.WriteLine("No QueryTables found in the worksheet.");
                }

                // Save the workbook (lifecycle rule)
                string outputPath = "QueryTableResultRangeDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            QueryTableResultRangeLogger.Run();
        }
    }
}

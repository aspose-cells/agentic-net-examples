// Title: How to obtain and log the ResultRange address of a QueryTable using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that checks a worksheet for QueryTables, retrieves the first QueryTable's ResultRange, and prints its address with Aspose.Cells. | Create a C# loop that iterates over all QueryTables in a worksheet, extracts each ResultRange, and writes the addresses to the console. | Add comprehensive error handling to a C# Aspose.Cells routine that safely accesses ResultRange even when no QueryTables exist, and logs an appropriate message.
// Common Searches: Aspose.Cells C# get address of QueryTable ResultRange | How to log QueryTable result range in a .NET Excel workbook | Check for QueryTables before accessing ResultRange with Aspose.Cells | Iterate through multiple QueryTables and retrieve their ResultRange addresses in C# | Aspose.Cells handling missing QueryTables when reading ResultRange
// Tags: Aspose.Cells QueryTable ResultRange address | C# retrieve QueryTable result range | enumerate QueryTables Aspose.Cells | log Excel query table range C# | error handling missing QueryTables Aspose.Cells

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Demonstrates creating a workbook, optionally populating data, checking for QueryTables, retrieving the first QueryTable's read‑only ResultRange, logging its address, and saving the file, with guidance for iterating multiple tables and handling absent QueryTables.
public class QueryTableResultRangeDemo
{
    public static void Run()
    {
        try
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data to simulate a query table
            worksheet.Cells["A1"].PutValue("ID");
            worksheet.Cells["B1"].PutValue("Name");
            worksheet.Cells["A2"].PutValue(1);
            worksheet.Cells["B2"].PutValue("John");
            worksheet.Cells["A3"].PutValue(2);
            worksheet.Cells["B3"].PutValue("Mary");

            // In a real scenario a QueryTable would be added via an external connection.
            // Here we simply check if any QueryTables exist and retrieve the ResultRange.
            if (worksheet.QueryTables.Count > 0)
            {
                // Get the first QueryTable
                QueryTable queryTable = worksheet.QueryTables[0];

                // Obtain the ResultRange (read‑only property)
                AsposeRange resultRange = queryTable.ResultRange;

                // Log the address for downstream processing
                Console.WriteLine("QueryTable ResultRange Address: " + resultRange.Address);
            }
            else
            {
                Console.WriteLine("No QueryTables found in the worksheet.");
            }

            // Save the workbook (save rule)
            workbook.Save("QueryTableResultRangeDemo.xlsx");
            Console.WriteLine("Workbook saved successfully.");
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
        QueryTableResultRangeDemo.Run();
    }
}

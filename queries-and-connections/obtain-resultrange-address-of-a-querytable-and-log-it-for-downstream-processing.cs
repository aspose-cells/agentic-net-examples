// Title: How to retrieve and log the ResultRange address of a QueryTable with Aspose.Cells for .NET (C#)
// AI Prompts: Write a C# function that opens an Excel file using Aspose.Cells, finds the first QueryTable on a worksheet, and returns its ResultRange address as a string. | Generate C# code to iterate through all QueryTables in a workbook with Aspose.Cells and print each ResultRange address to the console. | Create a reusable method in C# that accepts a workbook path and a QueryTable index, then logs the corresponding ResultRange address using Aspose.Cells.
// Common Searches: Aspose.Cells C# get address of QueryTable result range | How to print QueryTable ResultRange address in .NET Excel processing | Retrieve QueryTable output range address with Aspose.Cells for .NET
// Tags: Aspose.Cells QueryTable ResultRange retrieval | C# log QueryTable address with Aspose | Excel workbook QueryTable range extraction .NET | Iterate QueryTables Aspose.Cells

using System;
using Aspose.Cells;

// Loads an Excel workbook with Aspose.Cells, accesses the first worksheet, checks for QueryTables, obtains the ResultRange.Address of the first QueryTable, and writes the address to the console for downstream processing.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (or any specific worksheet by index/name)
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure the worksheet contains at least one QueryTable
        if (worksheet.QueryTables.Count > 0)
        {
            // Get the first QueryTable (adjust index if needed)
            QueryTable queryTable = worksheet.QueryTables[0];

            // Obtain the address of the ResultRange
            string resultRangeAddress = queryTable.ResultRange.Address;

            // Log the address for downstream processing
            Console.WriteLine("QueryTable ResultRange address: " + resultRangeAddress);
        }
        else
        {
            Console.WriteLine("No QueryTable found in the worksheet.");
        }
    }
}

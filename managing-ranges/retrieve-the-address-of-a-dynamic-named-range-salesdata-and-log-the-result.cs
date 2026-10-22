// Title: Retrieve the address of the dynamic named range "SalesData" from an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: how to obtain the RefersTo address of a named range with Aspose.Cells in C# | C# code to read the address of a dynamic named range from an .xlsx file using Aspose.Cells | example of using GetRangeByName to get a named range address in Aspose.Cells for .NET | log the full cell reference of a named range in an Excel workbook with Aspose.Cells | retrieve and display the sheet and cell range of a named range called SalesData in C#
// Tags: Aspose.Cells GetRangeByName RefersTo address | dynamic named range address extraction .NET | named range address logging C# | Excel workbook named range lookup Aspose.Cells | retrieve named range cell reference C# | SalesData named range address retrieval

using System;
using System.IO;
using Aspose.Cells;

// The sample loads 'input.xlsx', checks that the file exists, uses Workbook.Worksheets.GetRangeByName to locate the dynamic named range "SalesData", reads its RefersTo property to obtain the full address (e.g., Sheet1!A1:B10), and writes that address to the console while handling missing files, absent ranges, and other exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Ensure the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range "SalesData"
            // Use GetRangeByName to obtain the Aspose.Cells.Range object
            Aspose.Cells.Range salesRange = workbook.Worksheets.GetRangeByName("SalesData");

            if (salesRange == null)
            {
                Console.WriteLine("The named range \"SalesData\" does not exist in the workbook.");
                return;
            }

            // Get the address of the resolved range (e.g., Sheet1!A1:B10)
            string address = salesRange.RefersTo;

            // Log the result
            Console.WriteLine($"SalesData address: {address}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

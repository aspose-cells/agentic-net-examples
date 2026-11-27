// Title: Load an .xls workbook and enumerate its DataMashup queries using Aspose.Cells for .NET
// Common Searches: asp.net load .xls file and get DataMashup queries with Aspose.Cells | c# read DataMashup property from a legacy Excel workbook using Aspose.Cells | enumerate the queries collection inside DataMashup of an .xls workbook in .NET | handle cases where DataMashup is missing when opening an .xls file with Aspose.Cells | use dynamic object to access DataMashup in Aspose.Cells C# example
// Tags: open xls workbook Aspose.Cells | retrieve DataMashup queries .NET | dynamic DataMashup access in C# | handle absent DataMashup in workbook | count DataMashup queries Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample checks for the presence of an input.xls file, loads it into an Aspose.Cells Workbook, accesses the DataMashup property via a dynamic object, safely reads the Queries collection if available, outputs the number of queries, and includes robust exception handling for missing DataMashup or unexpected members.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xls";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"File not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Use dynamic to access DataMashup without compile‑time dependency
            dynamic dataMashup = workbook.DataMashup;

            if (dataMashup != null)
            {
                try
                {
                    var queries = dataMashup.Queries;
                    if (queries != null)
                    {
                        Console.WriteLine($"Number of queries in DataMashup: {queries.Count}");
                    }
                    else
                    {
                        Console.WriteLine("DataMashup has no queries.");
                    }
                }
                catch (Exception innerEx)
                {
                    // Handle cases where the DataMashup object does not expose expected members
                    Console.WriteLine($"Unable to read queries from DataMashup: {innerEx.Message}");
                }
            }
            else
            {
                Console.WriteLine("DataMashup is not available in this workbook.");
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

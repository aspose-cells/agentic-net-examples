// Title: Create a case‑insensitive header‑to‑column index dictionary by scanning the first row of an Aspose.Cells worksheet in C#
// AI Prompts: Write C# code that reads the first row of an Aspose.Cells worksheet and builds a Dictionary<string,int> where each key is a header text (ignoring case) and each value is the zero‑based column index. | Demonstrate how to query the dictionary for a specific header name and gracefully handle the situation when the header does not exist. | Adjust the iteration logic to stop when an empty cell is encountered in the header row and explain the benefit of early termination.
// Common Searches: aspocells c# get column index from header name in excel worksheet | how to build a header lookup dictionary with Aspose.Cells for .NET | case insensitive header mapping Aspose.Cells C# example | iterate over first row cells to create column index map using Aspose.Cells | retrieve column number by column header using Aspose.Cells API
// Tags: Aspose.Cells header dictionary | C# map Excel headers to column indexes | header lookup Aspose.Cells | first row iteration Aspose.Cells | column index mapping .NET

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHeaderMapping
{
    // The example creates a workbook, fills the first row with sample headers, then iterates the header cells up to the last populated column to construct a case‑insensitive Dictionary<string,int> that maps each header string to its column index, shows a lookup for a specific header, and saves the file.
    public class HeaderMapper
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample header row (row index 0)
                worksheet.Cells["A1"].PutValue("Name");
                worksheet.Cells["B1"].PutValue("Age");
                worksheet.Cells["C1"].PutValue("Country");
                worksheet.Cells["D1"].PutValue("Salary");

                // Dictionary to hold header name -> column index mapping
                Dictionary<string, int> headerToIndex = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                // Determine the last column that contains data in the header row
                int lastColumn = worksheet.Cells.MaxDataColumn;

                // Iterate through each column in the header row
                for (int col = 0; col <= lastColumn; col++)
                {
                    string header = worksheet.Cells[0, col].StringValue;
                    if (!string.IsNullOrEmpty(header))
                    {
                        // Add mapping to the dictionary
                        headerToIndex[header] = col;
                    }
                }

                // Demonstrate the lookup dictionary
                Console.WriteLine("Header to Column Index Mapping:");
                foreach (var kvp in headerToIndex)
                {
                    Console.WriteLine($"Header \"{kvp.Key}\" => Column Index {kvp.Value}");
                }

                // Example: retrieve column index for a specific header
                if (headerToIndex.TryGetValue("Country", out int countryCol))
                {
                    Console.WriteLine($"\nThe column index for \"Country\" is {countryCol}.");
                }

                // Save the workbook (optional, demonstrates lifecycle usage)
                string outputPath = "HeaderMapping.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }
                workbook.Save(outputPath);
                Console.WriteLine($"\nWorkbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            HeaderMapper.Run();
        }
    }
}

// Title: Append version number and author initials to an Excel table comment using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads a workbook with Aspose.Cells, retrieves the first ListObject, and appends a version string and author initials to its Comment property before saving the file. | Refactor an existing Aspose.Cells example to merge any existing table comment with new tracking metadata (version and author) and persist the updated workbook.
// Common Searches: how to add version info to an Excel table comment with Aspose.Cells C# | Aspose.Cells C# update ListObject comment with author initials | append tracking metadata to Excel table comment programmatically .NET | modify existing table comment without overwriting previous text Aspose.Cells | save workbook after changing Excel table comment using Aspose.Cells
// Tags: Aspose.Cells update ListObject comment | C# append version metadata to Excel table | Excel table comment tracking with Aspose | save workbook after comment modification Aspose.Cells | add author initials to ListObject comment .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExample
{
    // The example loads an XLSX workbook, checks for a table, builds version and author strings, concatenates them with any existing table comment, updates the ListObject's Comment property, and saves the workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Define input and output file paths
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one table (ListObject)
                if (worksheet.ListObjects.Count == 0)
                {
                    Console.WriteLine("No tables found in the worksheet.");
                    return;
                }

                // Retrieve the first table
                ListObject table = worksheet.ListObjects[0];

                // Prepare version and author information
                string versionInfo = "Version: 1.2";
                string authorInitials = "AB";

                // Combine existing comment with new tracking information
                string updatedComment = string.IsNullOrWhiteSpace(table.Comment)
                    ? $"{versionInfo} {authorInitials}"
                    : $"{table.Comment} {versionInfo} {authorInitials}";

                // Update the table's comment
                table.Comment = updatedComment.Trim();

                // Save the workbook with the updated comment
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

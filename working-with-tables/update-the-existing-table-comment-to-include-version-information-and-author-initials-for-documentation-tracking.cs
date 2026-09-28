// Title: Append version number and author initials to an Excel table comment using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an existing .xlsx file, retrieves the first ListObject on the first worksheet, and appends a version string and author initials to its Comment property with Aspose.Cells. | Create a C# snippet that checks for an existing table comment, concatenates version and author metadata, updates the ListObject.Comment, and saves the workbook.
// Common Searches: how to programmatically update a ListObject comment with version info using Aspose.Cells C# | Aspose.Cells add metadata to Excel table comment .NET example | C# append author initials to Excel table comment with Aspose.Cells library | update first table comment in workbook and save using Aspose.Cells for .NET
// Tags: Aspose.Cells modify ListObject comment | C# add version metadata to Excel table comment | Aspose.Cells update table comment .xlsx | track author initials in Excel table using Aspose.Cells | add version identifier to ListObject comment C#

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// Loads an existing workbook, accesses the first ListObject on the first worksheet, concatenates version and author initials to the table's Comment, and saves the updated file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Assume the target table is the first ListObject in the worksheet
            if (sheet.ListObjects.Count > 0)
            {
                ListObject table = sheet.ListObjects[0];

                // Retrieve the current comment (if any)
                string existingComment = table.Comment ?? string.Empty;

                // Define version and author information
                string versionInfo = "Version: 1.2";
                string authorInfo = "Author: AB";

                // Combine existing comment with new tracking info
                string newComment = $"{existingComment} {versionInfo}, {authorInfo}".Trim();

                // Update the table's comment
                table.Comment = newComment;
            }

            // Save the workbook with the updated comment
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

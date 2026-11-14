// Title: How to add version number and author initials to an Excel table (ListObject) comment using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add version and author to ListObject comment | update Excel table comment with version info using Aspose.Cells .NET | append author initials to Excel table comment programmatically Aspose.Cells | modify ListObject comment to include metadata in a C# workbook | Aspose.Cells example for editing table comment and saving workbook
// Tags: Aspose.Cells edit ListObject comment | C# add version metadata to Excel table | Aspose.Cells update table comment | Excel ListObject comment modification .NET | append author initials to Excel table using Aspose

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example loads an existing workbook, verifies the first worksheet contains a table (ListObject), reads the current table comment, appends a version string and author initials, updates the Comment property, and saves the workbook to a new file while handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";
        const string outputFile = "output.xlsx";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file \"{inputFile}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputFile);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one table (ListObject)
            if (worksheet.ListObjects.Count == 0)
            {
                Console.WriteLine("Error: No tables (ListObjects) found in the worksheet.");
                return;
            }

            // Access the first table on the worksheet
            ListObject table = worksheet.ListObjects[0];

            // Retrieve the current comment (may be empty)
            string currentComment = table.Comment ?? string.Empty;

            // Define version information and author initials
            string versionInfo = "Version: 1.2";
            string authorInfo = "Author: AB";

            // Append the new information to the existing comment
            table.Comment = $"{currentComment} {versionInfo}, {authorInfo}".Trim();

            // Save the modified workbook
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// Title: Delete a Named Range in Aspose.Cells (C#) and Verify Its Removal
// Description: This example creates a workbook, adds two named ranges (ObsoleteRange and ActiveRange), removes ObsoleteRange using NameCollection.Remove, confirms the deletion with GetRangeByName (returns null), ensures the remaining range is intact, and saves the file.
// Keywords: Aspose.Cells delete named range C# | remove defined name Aspose.Cells | NameCollection.Remove example | GetRangeByName null check | verify named range deletion
// Common Searches: how to delete a named range in Aspose.Cells .NET | check if a named range exists after removal Aspose.Cells | remove specific defined name without affecting others
// Developer Intent: Remove the "ObsoleteRange" named range from a workbook and confirm that it no longer exists while other named ranges stay unchanged.
// Use Cases: Clean up obsolete named ranges before publishing a spreadsheet. | Delete temporary ranges created during automated data processing. | Validate that a transformation script only removes intended named ranges.
// AI Prompts: Generate C# code using Aspose.Cells that deletes a named range and returns true if the operation succeeded. | Write a unit test that verifies a specific named range is removed and another remains after calling NameCollection.Remove. | Explain how NameCollection.Remove and GetRangeByName behave when the requested name is missing, including any exceptions or return values.

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    // This example creates a workbook, adds two named ranges (ObsoleteRange and ActiveRange), removes ObsoleteRange using NameCollection.Remove, confirms the deletion with GetRangeByName (returns null), ensures the remaining range is intact, and saves the file.
    public class DeleteNamedRangeDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add a named range that will be deleted later
                sheet.Cells.CreateRange("A1:B2").Name = "ObsoleteRange";

                // Add another named range to ensure only the target is removed
                sheet.Cells.CreateRange("C1:D2").Name = "ActiveRange";

                // Get the collection of defined names
                NameCollection names = workbook.Worksheets.Names;

                // Display count before removal
                Console.WriteLine("Named ranges count before removal: " + names.Count);

                // Remove the specific named range
                names.Remove("ObsoleteRange");

                // Verify removal by checking the count and attempting to retrieve the range
                Console.WriteLine("Named ranges count after removal: " + names.Count);

                // GetRangeByName returns null if the named range does not exist
                AsposeRange removedRange = workbook.Worksheets.GetRangeByName("ObsoleteRange");
                Console.WriteLine("ObsoleteRange exists after removal? " + (removedRange != null));

                // Also verify that the other named range still exists
                AsposeRange remainingRange = workbook.Worksheets.GetRangeByName("ActiveRange");
                Console.WriteLine("ActiveRange still exists? " + (remainingRange != null));

                // Save the workbook
                string outputPath = "DeleteNamedRangeDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine("Workbook saved to: " + Path.GetFullPath(outputPath));
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
            DeleteNamedRangeDemo.Run();
        }
    }
}

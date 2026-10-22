// Title: Read the address of a global named range on Sheet3 using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# get RefersTo address of a global named range on a specific sheet | How to access a workbook's global named range and retrieve its address with Aspose.Cells | Retrieve the address of a named range defined in Sheet3 using Aspose.Cells .NET API | C# code to read the address of a global named range from the Names collection in Aspose.Cells
// Tags: global named range address Aspose.Cells | Workbook.Names collection RefersTo C# | retrieve named range worksheet Aspose.Cells | read named range address .NET | Aspose.Cells named range on specific sheet

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook, accesses a global named range via the workbook's Worksheets.Names collection, verifies that it belongs to Sheet3, reads its address using the RefersTo property, prints the address, and saves the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string namedRange = "MyGlobalRange";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the global named range from the workbook's Names collection
            Name? globalName = null;
            try
            {
                // The Names collection provides indexed access by name; returns null if not found
                globalName = workbook.Worksheets.Names[namedRange];
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to retrieve named range '{namedRange}': {ex.Message}");
            }

            if (globalName != null)
            {
                // Obtain the actual cell range the name refers to
                Aspose.Cells.Range range = globalName.GetRange();

                // Verify that the range belongs to Sheet3
                if (range.Worksheet != null &&
                    range.Worksheet.Name.Equals("Sheet3", StringComparison.OrdinalIgnoreCase))
                {
                    // Display the address (e.g., Sheet3!$A$1:$B$5)
                    string address = globalName.RefersTo;
                    Console.WriteLine($"Global named range '{namedRange}' is on Sheet3 with address: {address}");
                }
                else
                {
                    Console.WriteLine($"Global named range '{namedRange}' does not refer to Sheet3.");
                }
            }
            else
            {
                Console.WriteLine($"Global named range '{namedRange}' was not found in the workbook.");
            }

            // Save the workbook (optional, if any modifications were made)
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

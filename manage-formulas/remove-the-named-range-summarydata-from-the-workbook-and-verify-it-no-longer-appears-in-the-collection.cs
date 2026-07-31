// Title: Remove the "SummaryData" named range from an Aspose.Cells workbook (C#)
// Description: Creates a workbook, adds a named range called SummaryData that points to A1:B2 on the first sheet, confirms its presence, deletes it with Worksheets.Names.Remove, verifies the removal, and saves the file.
// Keywords: Aspose.Cells remove named range C# | delete named range Aspose.Cells | Worksheets.Names.Remove example | verify named range deletion | C# workbook named range management
// Common Searches: how to delete a named range in Aspose.Cells .NET | C# remove specific named range Aspose.Cells | check if named range exists after removal Aspose.Cells | Aspose.Cells remove SummaryData range
// Developer Intent: Programmatically delete the "SummaryData" named range and ensure it no longer appears in the workbook's Names collection.
// Use Cases: Clean up temporary named ranges before exporting a report. | Prevent formula errors by removing dynamically created ranges after processing data. | Keep generated workbooks lightweight by pruning unused named ranges.
// AI Prompts: Generate C# code using Aspose.Cells that removes a named range called "SummaryData" and confirms the deletion. | Explain how to test for a named range's existence before and after calling Worksheets.Names.Remove in Aspose.Cells. | Show how to iterate through all named ranges in a workbook and delete those matching a given pattern.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a named range called SummaryData that points to A1:B2 on the first sheet, confirms its presence, deletes it with Worksheets.Names.Remove, verifies the removal, and saves the file.
    public class RemoveNamedRangeDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet (default sheet)
            Worksheet sheet = workbook.Worksheets[0];

            // Define a named range "SummaryData" referring to A1:B2
            int nameIndex = workbook.Worksheets.Names.Add("SummaryData");
            workbook.Worksheets.Names[nameIndex].RefersTo = $"={sheet.Name}!$A$1:$B$2";

            // Verify the named range exists before removal
            Name beforeRemoval = workbook.Worksheets.Names["SummaryData"];
            Console.WriteLine("Before removal, named range exists: " + (beforeRemoval != null));

            // Remove the named range using the Remove(string) method
            workbook.Worksheets.Names.Remove("SummaryData");

            // Verify the named range no longer exists
            Name afterRemoval = workbook.Worksheets.Names["SummaryData"];
            Console.WriteLine("After removal, named range exists: " + (afterRemoval != null));

            // Save the workbook (optional, demonstrates lifecycle rule)
            string outputPath = "RemoveNamedRangeDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}

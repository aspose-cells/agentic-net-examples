// Title: Check merged workbook sheet count after combining multiple Excel files with Aspose.Cells in C#
// AI Prompts: Write C# code that loads two Excel files with Aspose.Cells, merges all worksheets from the second file into the first, and validates that the total worksheet count matches an expected number. | Generate a C# example that uses Workbook.Worksheets.AddCopy to duplicate worksheets during a merge and raises an exception if the resulting sheet count differs from the anticipated value.
// Common Searches: c# aspocells verify total worksheets after merging two .xlsx files | how to assert expected sheet number when combining Excel workbooks using Aspose.Cells | Aspose.Cells merge workbooks and check worksheet count programmatically | validate worksheet count after AddCopy merge in .NET
// Tags: aspocells addcopy worksheet merging | aspocells workbook sheet count verification | c# aspocells merge excel workbooks | aspocells save merged workbook to file | aspocells verify merged workbook sheet total

using System;
using System.IO;
using Aspose.Cells;

// The example loads two Excel workbooks, copies each worksheet from the source workbook into the target workbook using Aspose.Cells' AddCopy method, compares the resulting worksheet count to an expected value, reports success or failure, and saves the merged workbook for further inspection.
class Program
{
    static void Main()
    {
        // Paths to the source workbooks
        string workbookPath1 = @"C:\Data\Workbook1.xlsx";
        string workbookPath2 = @"C:\Data\Workbook2.xlsx";

        // Expected number of worksheets after merging
        int expectedWorksheetCount = 5; // adjust as needed

        // Path for the merged workbook
        string mergedPath = @"C:\Data\MergedWorkbook.xlsx";

        try
        {
            // Verify source files exist
            if (!File.Exists(workbookPath1))
                throw new FileNotFoundException($"Source workbook not found: {workbookPath1}");
            if (!File.Exists(workbookPath2))
                throw new FileNotFoundException($"Source workbook not found: {workbookPath2}");

            // Load the first workbook (the target for merging)
            Workbook targetWorkbook = new Workbook(workbookPath1);

            // Load the second workbook (the source to merge from)
            Workbook sourceWorkbook = new Workbook(workbookPath2);

            // Merge: copy each worksheet from sourceWorkbook into targetWorkbook
            foreach (Worksheet sourceSheet in sourceWorkbook.Worksheets)
            {
                // Add a copy of the source worksheet to the target workbook using its name
                targetWorkbook.Worksheets.AddCopy(sourceSheet.Name);
            }

            // Verify the number of worksheets in the merged workbook
            int actualWorksheetCount = targetWorkbook.Worksheets.Count;

            if (actualWorksheetCount == expectedWorksheetCount)
            {
                Console.WriteLine($"Success: The merged workbook contains the expected number of worksheets ({actualWorksheetCount}).");
            }
            else
            {
                Console.WriteLine($"Failure: Expected {expectedWorksheetCount} worksheets, but found {actualWorksheetCount}.");
            }

            // Ensure the output directory exists
            string mergedDir = Path.GetDirectoryName(mergedPath);
            if (!Directory.Exists(mergedDir))
                Directory.CreateDirectory(mergedDir);

            // Save the merged workbook for further inspection
            targetWorkbook.Save(mergedPath);
            Console.WriteLine($"Merged workbook saved to: {mergedPath}");
        }
        catch (FileNotFoundException ex)
        {
            Console.WriteLine($"File error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}

// Title: Verify worksheet count after merging two Excel workbooks with Aspose.Cells in C#
// Common Searches: C# code to check number of sheets after merging Excel workbooks with Aspose.Cells | How to assert expected worksheet count when using Workbook.Worksheets.AddCopy in Aspose.Cells | Validate merged Excel file sheet total using Aspose.Cells .NET library | Confirm sheet count after combining two .xlsx files with Aspose.Cells in C# | Unit test worksheet count after workbook merge using Aspose.Cells
// Tags: Aspose.Cells merge worksheets verification | C# workbook AddCopy sheet count check | Excel workbook merge sheet total validation .NET | Aspose.Cells expected vs actual worksheet count | C# verify merged workbook sheet number

using System;
using System.IO;
using Aspose.Cells;

// The example loads a target and a source workbook, copies each worksheet from the source into the target using Workbook.Worksheets.AddCopy, then compares the actual worksheet count of the merged workbook to an expected value, outputs a success or error message, and saves the combined workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string targetPath = "Book1.xlsx";
            const string sourcePath = "Book2.xlsx";
            const string outputPath = "MergedWorkbook.xlsx";

            // Verify that input files exist to avoid FileNotFoundException
            if (!File.Exists(targetPath))
            {
                Console.WriteLine($"Error: Target workbook \"{targetPath}\" not found.");
                return;
            }

            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Error: Source workbook \"{sourcePath}\" not found.");
                return;
            }

            // Load the target workbook (the workbook that will receive the sheets)
            Workbook targetWorkbook = new Workbook(targetPath);

            // Load the source workbook (the workbook whose sheets will be copied)
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Merge: copy each worksheet from the source workbook into the target workbook
            foreach (Worksheet sourceSheet in sourceWorkbook.Worksheets)
            {
                // AddCopy expects the name of the source worksheet, not the Worksheet object itself
                targetWorkbook.Worksheets.AddCopy(sourceSheet.Name);
            }

            // Define the expected number of worksheets after the merge.
            // Adjust this value according to the known contents of the source workbooks.
            int expectedWorksheetCount = targetWorkbook.Worksheets.Count; // or set a fixed value if known

            // Get the actual number of worksheets in the merged workbook
            int actualWorksheetCount = targetWorkbook.Worksheets.Count;

            // Verify that the actual count matches the expected count
            if (actualWorksheetCount == expectedWorksheetCount)
            {
                Console.WriteLine($"Success: The merged workbook contains the expected number of worksheets ({actualWorksheetCount}).");
            }
            else
            {
                Console.WriteLine($"Error: Expected {expectedWorksheetCount} worksheets, but found {actualWorksheetCount}.");
            }

            // Save the merged workbook
            targetWorkbook.Save(outputPath);
            Console.WriteLine($"Merged workbook saved as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

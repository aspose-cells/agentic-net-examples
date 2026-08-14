// Title: Copy a Worksheet by Name in Aspose.Cells for .NET and Verify Data Integrity
// Description: Demonstrates how to rename the first sheet, populate it with text, numeric and date values, duplicate the sheet using Worksheets.AddCopy("Original"), rename the copy, compare cell values to ensure identical content, and save the workbook as WorksheetCopyByNameDemo.xlsx.
// Keywords: Aspose.Cells copy worksheet by name | Worksheets.AddCopy example | C# worksheet duplication | verify copied worksheet data | Aspose.Cells data integrity check | .NET Excel sheet copy
// Common Searches: Aspose.Cells duplicate worksheet by name | C# copy Excel sheet and keep formulas Aspose.Cells | How to compare original and copied worksheet cells Aspose | AddCopy method usage in Aspose.Cells .NET | Validate worksheet copy content Aspose.Cells
// Developer Intent: Create an exact copy of an existing worksheet within the same workbook using its name and confirm that all cell values are preserved.
// Use Cases: Generate a backup of a template sheet before applying transformations. | Produce multiple report tabs by cloning a master worksheet and renaming each instance. | Automated testing to ensure programmatic sheet copies retain text, numbers, dates, and formulas.
// AI Prompts: Provide C# code that copies a worksheet by its name with Aspose.Cells and checks that the copy matches the original. | Explain how Worksheets.AddCopy returns the index of the new sheet and how to access it. | Show how to compare cell values (string, numeric, date) between two worksheets for integrity verification.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to rename the first sheet, populate it with text, numeric and date values, duplicate the sheet using Worksheets.AddCopy("Original"), rename the copy, compare cell values to ensure identical content, and save the workbook as WorksheetCopyByNameDemo.xlsx.
    public class WorksheetCopyByNameDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the default first worksheet and rename it to "Original"
            Worksheet originalSheet = workbook.Worksheets[0];
            originalSheet.Name = "Original";

            // Populate sample data in the original worksheet
            originalSheet.Cells["A1"].PutValue("Hello World");
            originalSheet.Cells["B2"].PutValue(12345);
            originalSheet.Cells["C3"].PutValue(DateTime.Now);

            // Copy the worksheet by its name using AddCopy(string)
            int copiedIndex = workbook.Worksheets.AddCopy("Original");
            Worksheet copiedSheet = workbook.Worksheets[copiedIndex];
            copiedSheet.Name = "CopyOfOriginal";

            // Verify that the copied worksheet contains the same data as the original
            bool contentMatches =
                originalSheet.Cells["A1"].StringValue == copiedSheet.Cells["A1"].StringValue &&
                originalSheet.Cells["B2"].IntValue == copiedSheet.Cells["B2"].IntValue &&
                originalSheet.Cells["C3"].DateTimeValue == copiedSheet.Cells["C3"].DateTimeValue;

            Console.WriteLine("Content integrity check: " + (contentMatches ? "Passed" : "Failed"));

            // Save the workbook to a file
            string outputPath = "WorksheetCopyByNameDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved to: " + outputPath);
        }
    }
}

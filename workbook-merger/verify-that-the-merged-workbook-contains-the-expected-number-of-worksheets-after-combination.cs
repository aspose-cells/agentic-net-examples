// Title: Check worksheet count after merging workbooks with Aspose.Cells for .NET
// Description: Creates a source workbook with two worksheets and a destination workbook with one worksheet, calculates the expected total sheet count, merges the source into the destination using Workbook.Combine, and verifies that the resulting workbook contains the expected number of worksheets. The sample outputs the verification result and optionally saves the combined file.
// Keywords: Aspose.Cells | Workbook.Combine | C# | .NET | worksheet count verification | merge workbooks | combined workbook sheets | Excel automation | validate sheet number
// Common Searches: Aspose.Cells verify worksheet count after combine | C# count worksheets after Workbook.Combine | How to check number of sheets in merged Excel file using Aspose.Cells | Validate merged workbook sheet total .NET | Test workbook combine worksheet count
// Developer Intent: Confirm that Workbook.Combine adds all source worksheets to the destination, resulting in the correct total sheet count.
// Use Cases: Automated unit test for workbook merging logic | Pre‑deployment validation of Excel report generation pipelines | Runtime check before saving a combined workbook in a web service | Logging and alerting when worksheet count mismatch occurs | Batch processing of multiple workbooks where sheet count must meet business rules
// AI Prompts: Generate a reusable C# method that combines two Aspose.Cells Workbook objects and throws an exception if the resulting worksheet count does not equal the sum of the original counts. | Create an NUnit test that builds source and destination workbooks, calls Workbook.Combine, and asserts that the worksheet count matches the expected total. | Write code that logs detailed verification messages and saves the merged workbook only when the worksheet count validation succeeds.

using System;
using Aspose.Cells;

namespace VerifyCombinedWorkbook
{
    // Creates a source workbook with two worksheets and a destination workbook with one worksheet, calculates the expected total sheet count, merges the source into the destination using Workbook.Combine, and verifies that the resulting workbook contains the expected number of worksheets. The sample outputs the verification result and optionally saves the combined file.
    class Program
    {
        static void Main()
        {
            // Create source workbook with two worksheets
            Workbook sourceWorkbook = new Workbook();
            sourceWorkbook.Worksheets[0].Name = "SourceSheet1";
            sourceWorkbook.Worksheets.Add("SourceSheet2");

            // Create destination workbook with one worksheet
            Workbook destWorkbook = new Workbook();
            destWorkbook.Worksheets[0].Name = "DestSheet1";

            // Record expected worksheet count after combination
            int expectedCount = destWorkbook.Worksheets.Count + sourceWorkbook.Worksheets.Count;

            // Combine source workbook into destination workbook
            destWorkbook.Combine(sourceWorkbook);

            // Verify the worksheet count
            int actualCount = destWorkbook.Worksheets.Count;
            Console.WriteLine($"Expected worksheet count: {expectedCount}");
            Console.WriteLine($"Actual worksheet count:   {actualCount}");
            Console.WriteLine(actualCount == expectedCount
                ? "Verification succeeded: worksheet count matches expected value."
                : "Verification failed: worksheet count does not match expected value.");

            // Save the combined workbook (optional)
            destWorkbook.Save("CombinedWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}

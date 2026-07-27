// Title: C# – Verify Worksheet Count After Merging Workbooks with Aspose.Cells Combine
// Description: Creates a workbook with two sheets and another with three sheets, merges them using the Workbook.Combine method, and checks that the resulting file contains the expected five worksheets. The sample prints the verification result and optionally saves the combined file.
// Keywords: Aspose.Cells | C# | .NET | Workbook.Combine | merge workbooks | worksheet count | verify sheet total | combined workbook example
// Common Searches: Aspose.Cells combine workbooks sheet count | C# check number of worksheets after Combine | validate total worksheets in merged workbook Aspose.Cells | how to verify sheet count after workbook merge .NET
// Developer Intent: Confirm that the workbook produced by Combine contains the expected number of worksheets.
// Use Cases: Automated validation after consolidating multiple report files. | Unit‑test scenario to ensure no sheets are lost during a merge operation. | Pre‑save safety check that guarantees all source worksheets are present.
// AI Prompts: Generate a C# unit test using Aspose.Cells that creates two workbooks, combines them, and asserts Worksheets.Count equals the sum of source sheets. | Write code that merges three workbooks with Aspose.Cells, logs a warning if the final worksheet count differs from the calculated total, and handles exceptions gracefully. | Provide an example of using try‑catch to verify worksheet count after a Combine call and output a clear success or failure message.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook with two sheets and another with three sheets, merges them using the Workbook.Combine method, and checks that the resulting file contains the expected five worksheets. The sample prints the verification result and optionally saves the combined file.
    public class VerifyCombinedWorkbookWorksheetCount
    {
        public static void Run()
        {
            try
            {
                // Create the first workbook with two worksheets
                Workbook wb1 = new Workbook();
                wb1.Worksheets[0].Name = "FirstSheet1";
                wb1.Worksheets.Add("FirstSheet2");

                // Create the second workbook with three worksheets
                Workbook wb2 = new Workbook();
                wb2.Worksheets[0].Name = "SecondSheet1";
                wb2.Worksheets.Add("SecondSheet2");
                wb2.Worksheets.Add("SecondSheet3");

                // Combine the second workbook into the first workbook
                wb1.Combine(wb2);

                // Expected total worksheets after combination: 2 (from wb1) + 3 (from wb2) = 5
                int expectedCount = 5;
                int actualCount = wb1.Worksheets.Count;

                // Output verification results
                Console.WriteLine($"Expected worksheet count: {expectedCount}");
                Console.WriteLine($"Actual worksheet count after combine: {actualCount}");
                Console.WriteLine(actualCount == expectedCount ? "Verification passed." : "Verification failed.");

                // Save the combined workbook (optional)
                wb1.Save("CombinedWorkbook.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            VerifyCombinedWorkbookWorksheetCount.Run();
        }
    }
}

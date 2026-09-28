// Title: Verify worksheet count after combining two Excel workbooks with Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells Workbook.Combine to merge a source workbook into a destination workbook and assert that destWorkbook.Worksheets.Count equals sourceWorkbook.Worksheets.Count plus the original destination sheet count. | Generate C# code that creates two workbooks, combines them with Workbook.Combine, checks that the combined workbook contains the expected number of worksheets, and then saves the result.
// Common Searches: c# Aspose.Cells how to confirm total number of sheets after using Workbook.Combine | verify worksheet count after merging Excel files with Aspose.Cells .NET | expected vs actual sheet count after workbook combine in Aspose.Cells | Aspose.Cells combine workbooks and validate sheet total programmatically
// Tags: Aspose.Cells Workbook.Combine worksheet count verification | C# merge Excel workbooks validate sheet total | Aspose.Cells combined workbook save Xlsx | verify merged workbook sheet number .NET

using System;
using Aspose.Cells;

namespace VerifyCombinedWorkbook
{
    // The example creates a source workbook with two worksheets and a destination workbook with one worksheet, merges them using Workbook.Combine, validates that the resulting workbook contains the expected total number of worksheets, and saves the combined file as an XLSX document.
    class Program
    {
        static void Main()
        {
            // Create source workbook and add two worksheets
            Workbook sourceWorkbook = new Workbook();
            // Default workbook already has one worksheet at index 0
            sourceWorkbook.Worksheets[0].Name = "SourceSheet1";
            // Add a second worksheet
            sourceWorkbook.Worksheets.Add("SourceSheet2");

            // Create destination workbook and ensure it has one worksheet
            Workbook destWorkbook = new Workbook();
            destWorkbook.Worksheets[0].Name = "DestSheet1";

            // Expected total worksheets after combination
            int expectedWorksheetCount = sourceWorkbook.Worksheets.Count + destWorkbook.Worksheets.Count;

            // Combine the source workbook into the destination workbook
            destWorkbook.Combine(sourceWorkbook);

            // Verify the worksheet count
            int actualWorksheetCount = destWorkbook.Worksheets.Count;
            Console.WriteLine($"Expected worksheet count: {expectedWorksheetCount}");
            Console.WriteLine($"Actual worksheet count:   {actualWorksheetCount}");

            if (actualWorksheetCount == expectedWorksheetCount)
            {
                Console.WriteLine("Verification succeeded: worksheet count matches expected value.");
            }
            else
            {
                Console.WriteLine("Verification failed: worksheet count does not match expected value.");
            }

            // Save the combined workbook (optional, demonstrates usage of save)
            destWorkbook.Save("CombinedWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}

using System;
using Aspose.Cells;

class VerifyCombinedWorkbook
{
    static void Main()
    {
        // Create source workbook with two worksheets
        Workbook source = new Workbook();
        source.Worksheets[0].Name = "SourceSheet1";
        source.Worksheets.Add("SourceSheet2");

        // Create destination workbook with one worksheet
        Workbook dest = new Workbook();
        dest.Worksheets[0].Name = "DestSheet1";

        // Combine the source workbook into the destination workbook
        dest.Combine(source);

        // Expected worksheet count = original dest sheets + source sheets
        int expectedCount = 1 + source.Worksheets.Count; // dest originally had 1 sheet
        int actualCount = dest.Worksheets.Count;

        Console.WriteLine($"Expected worksheets: {expectedCount}");
        Console.WriteLine($"Actual worksheets after combine: {actualCount}");

        if (actualCount == expectedCount)
        {
            Console.WriteLine("Verification passed: worksheet count matches expected.");
        }
        else
        {
            Console.WriteLine("Verification failed: worksheet count does not match expected.");
        }

        // Save the combined workbook (optional)
        dest.Save("CombinedResult.xlsx", SaveFormat.Xlsx);
    }
}
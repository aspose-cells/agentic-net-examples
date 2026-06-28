using System;
using Aspose.Cells;

class VerifyCombinedWorkbook
{
    static void Main()
    {
        // Create source workbook with two worksheets
        Workbook source = new Workbook();
        source.Worksheets.Add("SourceSheet2"); // now the source has 2 sheets
        source.Worksheets[0].Cells["A1"].PutValue("Source Sheet 1");
        source.Worksheets[1].Cells["A1"].PutValue("Source Sheet 2");

        // Create destination workbook with one worksheet
        Workbook dest = new Workbook();
        dest.Worksheets[0].Name = "DestSheet1";
        dest.Worksheets[0].Cells["A1"].PutValue("Destination Sheet 1");

        // Combine the source workbook into the destination workbook
        dest.Combine(source);

        // Verify the number of worksheets after combination
        int expectedWorksheetCount = 3; // 1 original + 2 from source
        int actualWorksheetCount = dest.Worksheets.Count;

        Console.WriteLine($"Expected worksheet count: {expectedWorksheetCount}");
        Console.WriteLine($"Actual worksheet count:   {actualWorksheetCount}");
        Console.WriteLine(actualWorksheetCount == expectedWorksheetCount
            ? "Verification passed."
            : "Verification failed.");

        // Save the combined workbook (optional)
        dest.Save("CombinedResult.xlsx", SaveFormat.Xlsx);
    }
}
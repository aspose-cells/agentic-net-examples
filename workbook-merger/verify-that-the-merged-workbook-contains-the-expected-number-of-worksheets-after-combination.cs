using System;
using System.IO;
using Aspose.Cells;

namespace VerifyCombinedWorkbook
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create source workbook and add two worksheets
                Workbook sourceWorkbook = new Workbook();
                sourceWorkbook.Worksheets[0].Name = "SourceSheet1";

                // Add a second worksheet; Add returns a Worksheet object in recent Aspose.Cells versions
                Worksheet srcSheet2 = sourceWorkbook.Worksheets.Add("SourceSheet2");
                // (Optional) you can work with srcSheet2 here if needed

                // Create destination workbook and add one worksheet
                Workbook destWorkbook = new Workbook();
                destWorkbook.Worksheets[0].Name = "DestSheet";

                // Combine source into destination
                destWorkbook.Combine(sourceWorkbook);

                // Expected number of worksheets after combination:
                // Destination had 1, source had 2 => total should be 3
                int expectedCount = 3;
                int actualCount = destWorkbook.Worksheets.Count;

                // Verify and output the result
                Console.WriteLine($"Expected worksheet count: {expectedCount}");
                Console.WriteLine($"Actual worksheet count:   {actualCount}");
                Console.WriteLine(actualCount == expectedCount
                    ? "Verification succeeded: worksheet count matches expected."
                    : "Verification failed: worksheet count does not match expected.");

                // Optionally save the combined workbook (demonstrates lifecycle usage)
                string outputPath = "CombinedWorkbook.xlsx";
                destWorkbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Combined workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
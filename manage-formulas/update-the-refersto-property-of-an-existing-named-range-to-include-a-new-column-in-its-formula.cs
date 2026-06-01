using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    public class UpdateNamedRangeRefersTo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled error: {ex.Message}");
            }
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Sheet1";

                // Populate sample data in columns A and B (rows 1‑5)
                for (int i = 0; i < 5; i++)
                {
                    sheet.Cells[i, 0].PutValue($"A{i + 1}");
                    sheet.Cells[i, 1].PutValue($"B{i + 1}");
                }

                // Add a named range that refers to A1:B5
                int nameIndex = workbook.Worksheets.Names.Add("MyRange");
                Name myRange = workbook.Worksheets.Names[nameIndex];
                myRange.RefersTo = "=Sheet1!$A$1:$B$5";

                Console.WriteLine("Original RefersTo: " + myRange.RefersTo);

                // Update the RefersTo to include column C (A1:C5)
                string updatedRefersTo = "=Sheet1!$A$1:$C$5";
                myRange.RefersTo = updatedRefersTo;

                Console.WriteLine("Updated RefersTo: " + myRange.RefersTo);

                // Retrieve the updated range and display its address
                AsposeRange updatedRange = myRange.GetRange();
                Console.WriteLine("Updated range address: " + updatedRange.Address);

                // Save the workbook
                string outputPath = "UpdatedNamedRange.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Runtime error: {ex.Message}");
            }
        }
    }
}
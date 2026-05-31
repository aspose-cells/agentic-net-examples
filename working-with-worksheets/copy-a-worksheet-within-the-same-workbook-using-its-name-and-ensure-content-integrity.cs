using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class WorksheetCopyByNameDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook (in-memory)
            Workbook workbook = new Workbook();

            // Access the first worksheet and give it a meaningful name
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "OriginalSheet";

            // Populate some data in the source worksheet
            sourceSheet.Cells["A1"].PutValue("Header");
            sourceSheet.Cells["A2"].PutValue(123);
            sourceSheet.Cells["B2"].PutValue(456);
            sourceSheet.Cells["C2"].Formula = "=A2+B2";

            // Copy the worksheet within the same workbook using its name
            // AddCopy returns the index of the newly created worksheet
            int copiedIndex = workbook.Worksheets.AddCopy("OriginalSheet");
            Worksheet copiedSheet = workbook.Worksheets[copiedIndex];
            copiedSheet.Name = "CopiedSheet";

            // Verify that the copied content matches the source
            Console.WriteLine("Source A1: " + sourceSheet.Cells["A1"].StringValue);
            Console.WriteLine("Copied A1: " + copiedSheet.Cells["A1"].StringValue);

            Console.WriteLine("Source A2: " + sourceSheet.Cells["A2"].IntValue);
            Console.WriteLine("Copied A2: " + copiedSheet.Cells["A2"].IntValue);

            Console.WriteLine("Source B2: " + sourceSheet.Cells["B2"].IntValue);
            Console.WriteLine("Copied B2: " + copiedSheet.Cells["B2"].IntValue);

            Console.WriteLine("Source C2 Formula: " + sourceSheet.Cells["C2"].Formula);
            Console.WriteLine("Copied C2 Formula: " + copiedSheet.Cells["C2"].Formula);
            Console.WriteLine("Copied C2 Value (calculated): " + copiedSheet.Cells["C2"].IntValue);

            // Save the workbook (ensure the directory exists)
            string outputPath = "WorksheetCopyByNameDemo.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
    }
}
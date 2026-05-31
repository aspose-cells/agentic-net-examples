using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class WorksheetCopyByIndexDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (default contains one worksheet at index 0)
                Workbook workbook = new Workbook();

                // Access the original worksheet (index 0) and put some data
                Worksheet originalSheet = workbook.Worksheets[0];
                originalSheet.Name = "Original";
                originalSheet.Cells["A1"].PutValue("This is the original sheet");
                originalSheet.Cells["B2"].PutValue(12345);

                // Copy the worksheet using its numeric index (0)
                int copiedIndex = workbook.Worksheets.AddCopy(0); // returns the index of the new sheet
                Worksheet copiedSheet = workbook.Worksheets[copiedIndex];
                copiedSheet.Name = "Copied";

                // Verify that the copy contains the same data as the original
                string originalValue = originalSheet.Cells["A1"].StringValue;
                string copiedValue = copiedSheet.Cells["A1"].StringValue;
                double originalNumber = originalSheet.Cells["B2"].DoubleValue;
                double copiedNumber = copiedSheet.Cells["B2"].DoubleValue;

                Console.WriteLine("Verification of copied content:");
                Console.WriteLine($"A1 - Original: '{originalValue}' | Copied: '{copiedValue}'");
                Console.WriteLine($"B2 - Original: {originalNumber} | Copied: {copiedNumber}");

                // Modify the copied sheet to demonstrate that it is a separate instance
                copiedSheet.Cells["A1"].PutValue("This is the copied sheet");
                copiedSheet.Cells["B2"].PutValue(98765);

                // Show that the original sheet remains unchanged
                Console.WriteLine("\nAfter modifying the copied sheet:");
                Console.WriteLine($"Original A1: {originalSheet.Cells["A1"].StringValue}");
                Console.WriteLine($"Copied   A1: {copiedSheet.Cells["A1"].StringValue}");
                Console.WriteLine($"Original B2: {originalSheet.Cells["B2"].DoubleValue}");
                Console.WriteLine($"Copied   B2: {copiedSheet.Cells["B2"].DoubleValue}");

                // Save the workbook
                string outputPath = "WorksheetCopyByIndexDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"\nWorkbook saved as '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            WorksheetCopyByIndexDemo.Run();
        }
    }
}
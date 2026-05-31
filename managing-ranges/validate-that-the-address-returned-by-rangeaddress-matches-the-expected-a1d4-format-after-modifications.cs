using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeAddressValidation
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Create a range covering cells A1 to D4
                AsposeRange range = cells.CreateRange("A1:D4");

                // Verify that the Address property returns the expected A1:D4 string
                const string expectedAddress = "A1:D4";
                string actualAddress = range.Address;

                if (actualAddress != expectedAddress)
                {
                    Console.WriteLine($"Address mismatch. Expected: {expectedAddress}, Actual: {actualAddress}");
                }
                else
                {
                    Console.WriteLine($"Address correctly matches expected format: {actualAddress}");
                }

                // Modify the range by moving it one row down and one column right
                AsposeRange offsetRange = range.GetOffset(1, 1); // Should now be B2:E5

                // Verify the new address
                const string expectedOffsetAddress = "B2:E5";
                string actualOffsetAddress = offsetRange.Address;

                if (actualOffsetAddress != expectedOffsetAddress)
                {
                    Console.WriteLine($"Offset address mismatch. Expected: {expectedOffsetAddress}, Actual: {actualOffsetAddress}");
                }
                else
                {
                    Console.WriteLine($"Offset address correctly matches expected format: {actualOffsetAddress}");
                }

                // Determine output file path and ensure directory exists
                string outputPath = "RangeAddressValidation.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
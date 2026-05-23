using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
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

                // Create a range that starts at cell B2 (use fully‑qualified Aspose.Cells.Range to avoid ambiguity)
                Aspose.Cells.Range range = cells.CreateRange("B2");

                // Expand the range to the entire row that contains the original range
                Aspose.Cells.Range entireRow = range.EntireRow;

                // Retrieve the address of the entire row range
                string entireRowAddress = entireRow.Address;

                // Output the address to the console
                Console.WriteLine("Address of the entire row: " + entireRowAddress);

                // Save the workbook (optional, just to demonstrate lifecycle usage)
                string outputPath = "EntireRowAddressDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Runtime safety: report any errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
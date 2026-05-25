using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class ConcatenateRangeStrings
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate sample data in the range B2:D4
                cells["B2"].PutValue("Alpha");
                cells["C2"].PutValue(123);               // numeric value will be converted to string
                cells["D2"].PutValue(DateTime.Now);      // date will be formatted as string
                cells["B3"].PutValue("Beta");
                cells["C3"].PutValue("Gamma");
                cells["D3"].PutValue("Delta");
                cells["B4"].PutValue("Epsilon");
                cells["C4"].PutValue("Zeta");
                cells["D4"].PutValue("Eta");

                // Define the range to read
                string rangeAddress = "B2:D4";
                Aspose.Cells.Range range = cells.CreateRange(rangeAddress);

                // Concatenate raw string values from each cell in the range
                StringBuilder sb = new StringBuilder();
                foreach (Cell cell in range)
                {
                    sb.Append(cell.StringValue);
                }

                // Write the concatenated result to a summary cell (e.g., A1)
                cells["A1"].PutValue(sb.ToString());

                // Save the workbook
                string outputPath = "ConcatenatedRange.xlsx";
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
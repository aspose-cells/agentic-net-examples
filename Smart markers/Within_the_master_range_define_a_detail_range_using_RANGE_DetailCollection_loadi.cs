using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsDetailRangeExample
{
    class Program
    {
        static void Main()
        {
            const string inputPath = "MasterWorkbook.xlsx";
            const string outputPath = "MasterWorkbook_WithDetail.xlsx";

            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                // Create a new workbook with a single worksheet if the source file does not exist
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
            }

            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Define the address of the detail range within the master range.
            string detailStart = "B2"; // upper‑left cell of the detail range
            string detailEnd   = "D5"; // lower‑right cell of the detail range

            // Create the detail range
            AsposeRange detailRange = cells.CreateRange(detailStart, detailEnd);

            // Assign a name to the detail range
            detailRange.Name = "DetailCollection";

            // Optionally add the range to the worksheet's runtime range collection.
            cells.Ranges.Add(detailRange);

            // Populate the detail range with sample data
            for (int i = 0; i < detailRange.RowCount; i++)
            {
                for (int j = 0; j < detailRange.ColumnCount; j++)
                {
                    detailRange[i, j].PutValue($"R{i + 1}C{j + 1}");
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
        }
    }
}
using System;
using Aspose.Cells;

namespace AsposeCellsQueryTablePreserveFormatting
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through each worksheet in the workbook
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                // Iterate through each QueryTable in the current worksheet
                foreach (QueryTable queryTable in worksheet.QueryTables)
                {
                    // Enable preserving formatting when the query table is refreshed
                    queryTable.PreserveFormatting = true;
                }
            }

            // Save the modified workbook to a new file
            workbook.Save("output.xlsx");
        }
    }
}
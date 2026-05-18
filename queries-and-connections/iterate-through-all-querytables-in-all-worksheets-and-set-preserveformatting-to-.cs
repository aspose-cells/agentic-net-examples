using System;
using Aspose.Cells;

namespace AsposeCellsQueryTablePreserveFormatting
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through each worksheet in the workbook
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                // Iterate through each QueryTable in the current worksheet
                foreach (QueryTable queryTable in worksheet.QueryTables)
                {
                    // Set PreserveFormatting to true
                    queryTable.PreserveFormatting = true;
                }
            }

            // Save the modified workbook
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}
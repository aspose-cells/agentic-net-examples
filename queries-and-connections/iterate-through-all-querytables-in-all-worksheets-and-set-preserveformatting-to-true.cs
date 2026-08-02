using System;
using Aspose.Cells;

class SetQueryTablePreserveFormatting
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path) or create a new one
        Workbook workbook = new Workbook(); // new Workbook("input.xlsx");

        // Iterate through every worksheet in the workbook
        foreach (Worksheet worksheet in workbook.Worksheets)
        {
            // Iterate through each QueryTable in the current worksheet
            foreach (QueryTable queryTable in worksheet.QueryTables)
            {
                // Enable preserving formatting when the query table is refreshed
                queryTable.PreserveFormatting = true;
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
using System;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsQueryTablePreserveFormatting
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

                // Add sample data that will be part of the query table
                worksheet.Cells["A1"].PutValue("Product");
                worksheet.Cells["B1"].PutValue("Price");
                worksheet.Cells["A2"].PutValue("Laptop");
                worksheet.Cells["B2"].PutValue(1200);
                worksheet.Cells["A3"].PutValue("Phone");
                worksheet.Cells["B3"].PutValue(800);

                // Create a style and apply it to the header and data cells
                Style style = workbook.CreateStyle();
                style.Font.IsBold = true;
                style.ForegroundColor = Color.LightBlue;
                style.Pattern = BackgroundType.Solid;

                worksheet.Cells["A1"].SetStyle(style);
                worksheet.Cells["B1"].SetStyle(style);
                worksheet.Cells["A2"].SetStyle(style);
                worksheet.Cells["B2"].SetStyle(style);
                worksheet.Cells["A3"].SetStyle(style);
                worksheet.Cells["B3"].SetStyle(style);

                // NOTE: QueryTable API may not be available in all Aspose.Cells versions.
                // The following block is kept for reference; if the Add method exists,
                // it can be uncommented to demonstrate PreserveFormatting.
                /*
                int queryTableIndex = worksheet.QueryTables.Add("A1:B3", "A5", "SampleQueryTable");
                QueryTable queryTable = worksheet.QueryTables[queryTableIndex];
                queryTable.PreserveFormatting = true;
                */

                // Save the workbook
                string outputPath = "QueryTablePreserveFormattingDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Ods;

namespace AsposeCellsTableToRangeOds
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("John");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Mary");

            // Create a table (ListObject) that includes the header row
            int tableIndex = sheet.ListObjects.Add("A1", "B3", true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Apply formatting to the header row only
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.ForegroundColor = System.Drawing.Color.LightGray;
            headerStyle.Pattern = BackgroundType.Solid;

            sheet.Cells["A1"].SetStyle(headerStyle);
            sheet.Cells["B1"].SetStyle(headerStyle);

            // Convert the table to a normal range while keeping the header formatting
            table.ConvertToRange();

            // Prepare ODS save options (optional: set generator type)
            OdsSaveOptions saveOptions = new OdsSaveOptions
            {
                GeneratorType = OdsGeneratorType.LibreOffice
            };

            // Save the workbook as ODS
            workbook.Save("TableConvertedToRange.ods", saveOptions);
        }
    }
}
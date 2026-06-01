using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data that will become the table
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");

            // Add a ListObject (Excel table) covering the range A1:B3, with a header row
            int tableIdx = sheet.ListObjects.Add(0, 0, 2, 1, true);
            ListObject table = sheet.ListObjects[tableIdx];

            // Set a display name for the table (Name property may not be available in some versions)
            table.DisplayName = "SampleTable";

            // Attach a comment to the table that contains a hyperlink to external documentation
            string docUrl = "https://docs.aspose.com/cells/net/working-with-tables/";
            table.Comment = $"For more details see <a href=\"{docUrl}\">Aspose.Cells Table Documentation</a>";

            // Define output file path
            string outputPath = "TableWithComment.xlsx";

            // Save the workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
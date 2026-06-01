using System;
using Aspose.Cells;

class ExportWorkbookToCsv
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample date values to cells
        worksheet.Cells["A1"].PutValue(DateTime.Now);
        worksheet.Cells["A2"].PutValue(new DateTime(2023, 5, 15, 13, 45, 0));

        // Define ISO 8601 date format (e.g., 2023-05-15T13:45:00)
        Style isoDateStyle = workbook.CreateStyle();
        isoDateStyle.Custom = "yyyy-MM-ddTHH:mm:ss";

        // Apply the style to the range containing date cells
        StyleFlag flag = new StyleFlag();
        flag.NumberFormat = true; // Ensure only the number format is applied
        worksheet.Cells.CreateRange("A1:A2").ApplyStyle(isoDateStyle, flag);

        // Save the workbook as CSV; dates will be output using the ISO 8601 format
        workbook.Save("ExportedDates.csv", SaveFormat.Csv);
    }
}
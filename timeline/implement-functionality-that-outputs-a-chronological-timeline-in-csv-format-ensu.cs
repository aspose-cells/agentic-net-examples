using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsTimelineCsvDemo
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // 2. Populate worksheet with chronological data (Date and Value)
            // Header row
            cells[0, 0].PutValue("Date");
            cells[0, 1].PutValue("Value");

            // Sample data rows (ensure dates are in chronological order)
            cells[1, 0].PutValue(new DateTime(2021, 2, 5));
            cells[1, 1].PutValue(50);

            cells[2, 0].PutValue(new DateTime(2022, 3, 8));
            cells[2, 1].PutValue(60);

            cells[3, 0].PutValue(new DateTime(2023, 4, 10));
            cells[3, 1].PutValue(70);

            cells[4, 0].PutValue(new DateTime(2024, 5, 16));
            cells[4, 1].PutValue(80);

            // Optional: Apply a date format to the Date column
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "yyyy-MM-dd";
            cells[1, 0].SetStyle(dateStyle);
            cells[2, 0].SetStyle(dateStyle);
            cells[3, 0].SetStyle(dateStyle);
            cells[4, 0].SetStyle(dateStyle);

            // 3. Save the worksheet as CSV with proper delimiter and header handling
            TxtSaveOptions saveOptions = new TxtSaveOptions
            {
                // Use comma as the separator for CSV
                Separator = ',',
                // Ensure that blank rows (if any) still output separators
                KeepSeparatorsForBlankRow = true,
                // Preserve the header row as is
                Encoding = Encoding.UTF8
            };

            // Save to a memory stream to obtain the CSV content as a string
            using (MemoryStream csvStream = new MemoryStream())
            {
                workbook.Save(csvStream, saveOptions);
                string csvContent = Encoding.UTF8.GetString(csvStream.ToArray());

                // Output the CSV content to console (or further processing)
                Console.WriteLine(csvContent);
            }
        }
    }
}
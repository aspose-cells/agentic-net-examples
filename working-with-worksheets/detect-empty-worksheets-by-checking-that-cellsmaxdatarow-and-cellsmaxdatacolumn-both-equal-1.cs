using System;
using Aspose.Cells;

namespace AsposeCellsEmptyWorksheetDetection
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook or load an existing one
            // For demonstration, we create a new workbook with three sheets
            Workbook workbook = new Workbook();

            // Add two more worksheets
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");

            // Populate data only in the first worksheet
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Cells["A1"].PutValue("Sample Data");

            // The second and third worksheets remain completely empty

            // Iterate through all worksheets to detect empty ones
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];
                // A worksheet is considered empty when both MaxDataRow and MaxDataColumn are -1
                if (sheet.Cells.MaxDataRow == -1 && sheet.Cells.MaxDataColumn == -1)
                {
                    Console.WriteLine($"Worksheet \"{sheet.Name}\" is empty.");
                }
                else
                {
                    Console.WriteLine($"Worksheet \"{sheet.Name}\" contains data.");
                }
            }

            // Optionally, save the workbook (not required for detection)
            // workbook.Save("DetectedEmptySheets.xlsx", SaveFormat.Xlsx);
        }
    }
}
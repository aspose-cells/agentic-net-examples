using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    public class UpdateNamedRangeWithCurrentDate
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet and set its name
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Sheet1";

                // Put the current date into cell B2
                sheet.Cells["B2"].PutValue(DateTime.Now);

                // Add a named range that refers to cell B2
                int nameIndex = workbook.Worksheets.Names.Add("TodayDate");
                Name todayName = workbook.Worksheets.Names[nameIndex];
                todayName.RefersTo = "=Sheet1!$B$2";

                // Retrieve the range via the named range and update it with the current date
                AsposeRange dateRange = todayName.GetRange();
                if (dateRange != null && dateRange.RowCount > 0 && dateRange.ColumnCount > 0)
                {
                    dateRange[0, 0].PutValue(DateTime.Now);
                }

                // Define output file path
                string outputPath = "UpdatedNamedRange.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
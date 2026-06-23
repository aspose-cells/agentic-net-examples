using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    public class AutoFilterColumnAJ
    {
        public static void Run()
        {
            try
            {
                // Load an existing template if it exists; otherwise create a new workbook
                string templatePath = "Template.xlsx";
                Workbook workbook = File.Exists(templatePath) ? new Workbook(templatePath) : new Workbook();

                Worksheet sheet = workbook.Worksheets[0];

                // Determine the last row that contains data
                int lastRow = sheet.Cells.MaxDataRow;

                // Apply AutoFilter to column AJ (index 35) from header row (0) to the last data row
                sheet.AutoFilter.SetRange(0, 35, lastRow);

                // Show only rows where the value in column AJ is greater than 500
                sheet.AutoFilter.Custom(35, FilterOperatorType.GreaterThan, 500);

                // Apply the filter
                sheet.AutoFilter.Refresh();

                // Save the filtered workbook
                string outputPath = "FilteredColumnAJ.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            AutoFilterColumnAJ.Run();
        }
    }
}
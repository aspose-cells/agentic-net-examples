using System;
using Aspose.Cells;

namespace AsposeCellsDataFiltering
{
    class Program
    {
        static void Main(string[] args)
        {
            FilterXlsxDemo.Run();
        }
    }

    public class FilterXlsxDemo
    {
        public static void Run()
        {
            string sourcePath = "input.xlsx";

            LoadOptions loadOptions = new LoadOptions();
            loadOptions.AutoFilter = true;

            Workbook workbook = new Workbook(sourcePath, loadOptions);

            Worksheet sheet = workbook.Worksheets[0];

            sheet.AutoFilter.Range = "A1:C10";

            sheet.AutoFilter.AddFilter(1, "Electronics");

            sheet.AutoFilter.Refresh();

            string outputPath = "filtered_output.xlsx";
            workbook.Save(outputPath);
        }
    }
}
using System;
using Aspose.Cells;

namespace AsposeCellsDateFormattingDemo
{
    class Program
    {
        static void Main()
        {
            string inputPath = "input.xlsx";
            string outputPath = "output_formatted.xlsx";

            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            Workbook workbook = new Workbook(inputPath, loadOptions);
            Worksheet worksheet = workbook.Worksheets[0];

            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "yyyy-MM-dd";

            Aspose.Cells.Range usedRange = worksheet.Cells.MaxDisplayRange;
            foreach (Cell cell in usedRange)
            {
                if (cell.Type == CellValueType.IsDateTime)
                {
                    StyleFlag flag = new StyleFlag();
                    flag.NumberFormat = true;
                    cell.SetStyle(dateStyle, flag);
                }
            }

            workbook.Save(outputPath);
        }
    }
}
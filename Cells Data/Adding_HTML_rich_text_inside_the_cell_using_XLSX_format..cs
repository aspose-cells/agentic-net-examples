using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class HtmlStringInCellDemo
    {
        public static void Run()
        {
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;
            cells["A1"].HtmlString = "This is <b>bold</b> and <i>italic</i> text<br>with a line break";
            workbook.Save("HtmlStringDemo.xlsx", SaveFormat.Xlsx);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            HtmlStringInCellDemo.Run();
        }
    }
}
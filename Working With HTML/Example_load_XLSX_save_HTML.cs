using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class LoadXlsxSaveHtmlDemo
    {
        public static void Run()
        {
            // Path to the source XLSX file
            string sourcePath = "input.xlsx";

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(sourcePath);

            // Save the workbook as an HTML file
            workbook.Save("output.html", SaveFormat.Html);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            LoadXlsxSaveHtmlDemo.Run();
        }
    }
}
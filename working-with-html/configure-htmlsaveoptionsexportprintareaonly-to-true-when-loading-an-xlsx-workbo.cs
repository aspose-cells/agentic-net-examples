using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ExportPrintAreaOnlyDemo
    {
        public static void Run()
        {
            // Load an existing XLSX workbook from file
            Workbook workbook = new Workbook("input.xlsx");

            // Create HTML save options and enable exporting only the print area
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.ExportPrintAreaOnly = true; // Export only the defined print area

            // Save the workbook as an HTML file using the configured options
            workbook.Save("output.html", htmlOptions);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExportPrintAreaOnlyDemo.Run();
        }
    }
}
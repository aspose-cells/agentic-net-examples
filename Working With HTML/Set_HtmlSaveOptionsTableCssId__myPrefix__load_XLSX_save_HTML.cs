using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main()
        {
            HtmlSaveOptionsTableCssIdDemo.Run();
        }
    }

    public class HtmlSaveOptionsTableCssIdDemo
    {
        public static void Run()
        {
            // Load an existing XLSX workbook from disk
            Workbook workbook = new Workbook("input.xlsx");

            // Create HTML save options and set the TableCssId prefix
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
            saveOptions.TableCssId = "myPrefix_";

            // Save the workbook as an HTML file using the configured options
            workbook.Save("output.html", saveOptions);

            Console.WriteLine("Workbook saved to HTML with TableCssId set to \"myPrefix_\".");
        }
    }
}
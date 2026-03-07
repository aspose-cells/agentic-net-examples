using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class PageBreakPreviewDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Enable Page Break Preview mode
            worksheet.IsPageBreakPreview = true;

            // Set zoom level (optional, for demonstration)
            worksheet.Zoom = 100;

            // Output current settings
            Console.WriteLine("IsPageBreakPreview: " + worksheet.IsPageBreakPreview);
            Console.WriteLine("Zoom: " + worksheet.Zoom);

            // Save the workbook
            workbook.Save("PageBreakPreviewDemo_output.xlsx");
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            PageBreakPreviewDemo.Run();
        }
    }
}
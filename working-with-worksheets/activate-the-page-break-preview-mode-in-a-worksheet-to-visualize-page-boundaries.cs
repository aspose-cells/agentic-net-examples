using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ActivatePageBreakPreview
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Activate page break preview mode
            worksheet.IsPageBreakPreview = true;

            // Set zoom level (optional)
            worksheet.Zoom = 100;

            // Save the workbook
            workbook.Save("PageBreakPreviewDemo.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ActivatePageBreakPreview.Run();
        }
    }
}
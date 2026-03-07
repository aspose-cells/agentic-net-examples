using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class EnablePageBreakPreview
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Enable Page Break Preview mode
            worksheet.IsPageBreakPreview = true;

            // Alternatively, set the view type to PageBreakPreview
            worksheet.ViewType = ViewType.PageBreakPreview;

            // Set zoom to 100% for demonstration purposes
            worksheet.Zoom = 100;

            // Save the workbook
            workbook.Save("EnablePageBreakPreview_output.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            EnablePageBreakPreview.Run();
        }
    }
}
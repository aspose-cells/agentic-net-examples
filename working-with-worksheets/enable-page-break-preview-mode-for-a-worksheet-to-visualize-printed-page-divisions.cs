using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Author: Aspose.Cells .NET example
    class EnablePageBreakPreview
    {
        static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Enable Page Break Preview mode
            worksheet.IsPageBreakPreview = true;

            // Optional: set zoom to 100% for clearer view
            worksheet.Zoom = 100;

            // Save the workbook (save rule)
            workbook.Save("PageBreakPreviewDemo.xlsx");

            Console.WriteLine("Page Break Preview enabled. Workbook saved as PageBreakPreviewDemo.xlsx");
        }
    }
}
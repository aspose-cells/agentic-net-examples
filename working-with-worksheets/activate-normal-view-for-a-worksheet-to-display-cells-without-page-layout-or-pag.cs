using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ActivateNormalViewDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set the worksheet to Normal view (no page break or layout preview)
            worksheet.ViewType = ViewType.NormalView;
            // Ensure the page break preview flag is disabled
            worksheet.IsPageBreakPreview = false;

            // Save the workbook
            workbook.Save("NormalViewDemo.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ActivateNormalViewDemo.Run();
        }
    }
}
using System;
using Aspose.Cells;

namespace AsposeCellsViewModeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data to visualize the view modes
            sheet.Cells["A1"].PutValue("View Mode Demo");
            sheet.Cells["A2"].PutValue("Normal View");
            sheet.Cells["A3"].PutValue("Page Break Preview");
            sheet.Cells["A4"].PutValue("Page Layout View");

            // ---------- Normal View ----------
            // Set the worksheet view type to Normal (rule: Worksheet.ViewType)
            sheet.ViewType = ViewType.NormalView;
            // Save the workbook showing Normal view
            workbook.Save("Workbook_NormalView.xlsx");

            // ---------- Page Break Preview ----------
            // Switch to Page Break Preview mode
            sheet.ViewType = ViewType.PageBreakPreview;
            // Optionally enable the ruler for better visibility in preview
            sheet.IsRulerVisible = true;
            // Save the workbook showing Page Break Preview
            workbook.Save("Workbook_PageBreakPreview.xlsx");

            // ---------- Page Layout View ----------
            // Switch to Page Layout View mode
            sheet.ViewType = ViewType.PageLayoutView;
            // Ensure the ruler is visible in this view as well
            sheet.IsRulerVisible = true;
            // Save the workbook showing Page Layout view
            workbook.Save("Workbook_PageLayoutView.xlsx");

            // Cleanup
            workbook.Dispose();

            Console.WriteLine("Workbook view mode switching completed.");
        }
    }
}
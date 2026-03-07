using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ShowHideGridlinesAndHeadersDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // Worksheet 1 – Show gridlines and row/column headers
            // -------------------------------------------------
            Worksheet sheetShow = workbook.Worksheets[0]; // first worksheet
            sheetShow.Name = "ShowGridlines";

            // Make gridlines visible
            sheetShow.IsGridlinesVisible = true;

            // Make row and column headers visible
            sheetShow.IsRowColumnHeadersVisible = true;

            // Add some sample data
            sheetShow.Cells["A1"].PutValue("Gridlines: Visible");
            sheetShow.Cells["A2"].PutValue("Headers: Visible");

            // -------------------------------------------------
            // Worksheet 2 – Hide gridlines and row/column headers
            // -------------------------------------------------
            int sheetIndex = workbook.Worksheets.Add();
            Worksheet sheetHide = workbook.Worksheets[sheetIndex];
            sheetHide.Name = "HideGridlines";

            // Hide gridlines
            sheetHide.IsGridlinesVisible = false;

            // Hide row and column headers
            sheetHide.IsRowColumnHeadersVisible = false;

            // Add some sample data
            sheetHide.Cells["A1"].PutValue("Gridlines: Hidden");
            sheetHide.Cells["A2"].PutValue("Headers: Hidden");

            // Save the workbook to a file
            workbook.Save("ShowHideGridlinesAndHeaders.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ShowHideGridlinesAndHeadersDemo.Run();
        }
    }
}
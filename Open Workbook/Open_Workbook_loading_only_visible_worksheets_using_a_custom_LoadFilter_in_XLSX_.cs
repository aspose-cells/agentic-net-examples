using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a workbook with one visible and one hidden worksheet
        Workbook wb = new Workbook();

        // Configure visible sheet (default sheet)
        Worksheet visible = wb.Worksheets[0];
        visible.Name = "VisibleSheet";
        visible.Cells["A1"].PutValue("This will be loaded");

        // Add a hidden sheet
        Worksheet hidden = wb.Worksheets.Add("HiddenSheet");
        hidden.IsVisible = false;
        hidden.Cells["A1"].PutValue("This should not be loaded");

        // Save the workbook (optional)
        wb.Save("output.xlsx");
    }
}
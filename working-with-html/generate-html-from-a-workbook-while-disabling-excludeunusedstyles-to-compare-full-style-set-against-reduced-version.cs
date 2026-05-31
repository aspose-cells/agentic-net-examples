using System;
using System.Drawing;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample data with distinct styles
        Cell cell1 = workbook.Worksheets[0].Cells["A1"];
        cell1.PutValue("Styled text");
        Style style1 = workbook.CreateStyle();
        style1.Font.Color = Color.Red;
        style1.Font.IsBold = true;
        cell1.SetStyle(style1);

        Cell cell2 = workbook.Worksheets[0].Cells["A2"];
        cell2.PutValue("Different style");
        Style style2 = workbook.CreateStyle();
        style2.Font.Name = "Times New Roman";
        style2.Font.Size = 14;
        cell2.SetStyle(style2);

        // Set HTML save options to keep all styles (disable exclusion of unused styles)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExcludeUnusedStyles = false; // retain full style set in the generated HTML

        // Save the workbook as HTML using the configured options
        workbook.Save("FullStyles.html", htmlOptions);
    }
}
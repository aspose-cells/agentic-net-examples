using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main()
        {
            OxpsSaveDemo.Run();
        }
    }

    public class OxpsSaveDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data to the worksheet
            sheet.Cells["A1"].PutValue("Aspose.Cells OXPS Demo");
            sheet.Cells["A2"].PutValue(DateTime.Now);
            sheet.Cells["A3"].PutValue(12345);

            // Apply simple formatting to the header cell
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.Font.Color = Color.Blue;
            sheet.Cells["A1"].SetStyle(headerStyle);

            // Save the workbook as XPS (compatible with OXPS)
            workbook.Save("OxpsDemo.xps", SaveFormat.Xps);
        }
    }
}
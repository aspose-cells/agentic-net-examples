using System;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsExamples
{
    public class EnableNormalViewDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set the worksheet view type to Normal view
            worksheet.ViewType = ViewType.NormalView;

            // Verify the view type
            Console.WriteLine("Worksheet ViewType set to: " + worksheet.ViewType);

            // Save the workbook as an Excel file (default view will be normal)
            workbook.Save("NormalViewDemo.xlsx");

            // If exporting to DOCX and you want the output to be in normal view,
            // set the AsNormalView property on DocxSaveOptions
            DocxSaveOptions docxOptions = new DocxSaveOptions
            {
                AsNormalView = true
            };

            // Save the workbook as DOCX with normal view enabled
            workbook.Save("NormalViewDemo.docx", docxOptions);

            Console.WriteLine("Files saved successfully with Normal view enabled.");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            EnableNormalViewDemo.Run();
        }
    }
}
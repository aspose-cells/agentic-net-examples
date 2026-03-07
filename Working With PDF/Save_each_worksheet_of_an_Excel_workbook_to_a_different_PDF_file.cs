using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    public class SaveEachWorksheetToSeparatePdf
    {
        public static void Main()
        {
            // Create a new workbook and add sample data to multiple worksheets
            Workbook workbook = new Workbook();

            // First worksheet (default)
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Summary";
            sheet1.Cells["A1"].PutValue("Report Summary");
            sheet1.Cells["A2"].PutValue(DateTime.Now.ToString());

            // Add a second worksheet
            Worksheet sheet2 = workbook.Worksheets.Add("Data");
            for (int i = 0; i < 10; i++)
            {
                sheet2.Cells[i, 0].PutValue($"Item {i + 1}");
                sheet2.Cells[i, 1].PutValue(i * 10);
            }

            // Add a third worksheet
            Worksheet sheet3 = workbook.Worksheets.Add("Chart");
            sheet3.Cells["A1"].PutValue("Category");
            sheet3.Cells["B1"].PutValue("Value");
            sheet3.Cells["A2"].PutValue("A");
            sheet3.Cells["B2"].PutValue(30);
            sheet3.Cells["A3"].PutValue("B");
            sheet3.Cells["B3"].PutValue(45);
            sheet3.Cells["A4"].PutValue("C");
            sheet3.Cells["B4"].PutValue(25);

            // Loop through each worksheet and save it as an individual PDF file
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                // Create PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Restrict the save operation to the current worksheet only
                pdfOptions.SheetSet = new SheetSet(new int[] { i });

                // Build the output file name using the worksheet name
                string pdfFileName = $"{workbook.Worksheets[i].Name}.pdf";

                // Save the workbook (only the selected sheet) to PDF
                workbook.Save(pdfFileName, pdfOptions);
            }

            Console.WriteLine("All worksheets have been saved as separate PDF files.");
        }
    }
}
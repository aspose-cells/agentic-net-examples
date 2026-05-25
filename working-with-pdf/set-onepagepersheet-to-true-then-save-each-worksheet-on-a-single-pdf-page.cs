using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a workbook and add a few worksheets with sample data
        Workbook workbook = new Workbook();
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "FirstSheet";
        sheet1.Cells["A1"].PutValue("Data in first sheet");

        Worksheet sheet2 = workbook.Worksheets.Add("SecondSheet");
        sheet2.Cells["A1"].PutValue("Data in second sheet");

        Worksheet sheet3 = workbook.Worksheets.Add("ThirdSheet");
        sheet3.Cells["A1"].PutValue("Data in third sheet");

        // Configure PDF save options: each sheet will be rendered on a single PDF page
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.OnePagePerSheet = true; // Ensure one page per sheet

        // Iterate through all worksheets and save each one as an individual PDF file
        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            // Restrict rendering to the current worksheet only
            pdfOptions.SheetSet = new SheetSet(new int[] { i });

            // Build output file name using the worksheet name
            string outputPath = $"{workbook.Worksheets[i].Name}.pdf";

            // Save the workbook (only the selected sheet) to PDF
            workbook.Save(outputPath, pdfOptions);
        }
    }
}
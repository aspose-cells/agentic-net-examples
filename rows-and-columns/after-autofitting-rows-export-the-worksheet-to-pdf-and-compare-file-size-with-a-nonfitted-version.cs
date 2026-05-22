using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    public class AutoFitRowsPdfComparison
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data that will cause rows to need auto‑fitting
            sheet.Cells["A1"].PutValue("This is a very long text that should cause the row height to increase when AutoFitRows is applied.");
            sheet.Cells["A2"].PutValue("Short text");
            sheet.Cells["A3"].PutValue("Another long text with line breaks\nto demonstrate multi‑line row height adjustment.");

            // Enable text wrapping so row height is affected
            Style wrapStyle = sheet.Cells["A1"].GetStyle();
            wrapStyle.IsTextWrapped = true;
            sheet.Cells["A1"].SetStyle(wrapStyle);
            sheet.Cells["A3"].SetStyle(wrapStyle);

            // Define file paths
            string outputDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "AsposeRowsPdf");
            Directory.CreateDirectory(outputDir);
            string pdfWithoutFit = Path.Combine(outputDir, "WithoutAutoFit.pdf");
            string pdfWithFit = Path.Combine(outputDir, "WithAutoFit.pdf");

            // Save PDF before auto‑fitting rows
            workbook.Save(pdfWithoutFit, SaveFormat.Pdf);

            // Auto‑fit all rows in the worksheet
            sheet.AutoFitRows();

            // Save PDF after auto‑fitting rows
            workbook.Save(pdfWithFit, SaveFormat.Pdf);

            // Compare file sizes
            long sizeWithoutFit = new FileInfo(pdfWithoutFit).Length;
            long sizeWithFit = new FileInfo(pdfWithFit).Length;

            Console.WriteLine($"PDF size without AutoFitRows: {sizeWithoutFit} bytes");
            Console.WriteLine($"PDF size with AutoFitRows   : {sizeWithFit} bytes");

            if (sizeWithFit > sizeWithoutFit)
                Console.WriteLine("AutoFitRows increased the PDF file size.");
            else if (sizeWithFit < sizeWithoutFit)
                Console.WriteLine("AutoFitRows decreased the PDF file size.");
            else
                Console.WriteLine("PDF file size unchanged after AutoFitRows.");
        }
    }
}
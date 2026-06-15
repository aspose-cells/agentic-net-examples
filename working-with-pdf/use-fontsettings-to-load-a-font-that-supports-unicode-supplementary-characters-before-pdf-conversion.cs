using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Specify the folder that contains the Unicode‑supporting font files.
        // The second argument (true) tells Aspose.Cells to scan subfolders as well.
        FontConfigs.SetFontFolder("Fonts", true);

        // Create a new workbook and add some text that includes Unicode supplementary characters.
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Unicode test: 𝔘𝔫𝔦𝔠𝔬𝔡𝔢 🌟 漢字");

        // Set the workbook's default font to the custom font (replace with the actual font name present in the Fonts folder).
        workbook.DefaultStyle.Font.Name = "NotoSans";

        // Configure PDF save options to use the custom default font and ensure proper encoding.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            DefaultFont = "NotoSans",          // Font name that supports the characters.
            CheckWorkbookDefaultFont = true,   // Try workbook's default font first.
            FontEncoding = PdfFontEncoding.Identity
        };

        // Save the workbook as a PDF using the configured options.
        workbook.Save("output.pdf", pdfOptions);
    }
}
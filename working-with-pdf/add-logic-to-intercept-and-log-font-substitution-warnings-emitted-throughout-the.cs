using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

public class RenderingWarningCallback : IWarningCallback
{
    public int WarningCount { get; private set; }

    public void Warning(WarningInfo warningInfo)
    {
        // Log only font substitution warnings
        if (warningInfo.WarningType == WarningType.FontSubstitution)
        {
            Console.WriteLine($"Font substitution warning: {warningInfo.Description}");
            WarningCount++;
        }
    }
}

public class PdfConversionWithWarningLogging
{
    public static void Main()
    {
        // Create a workbook with a cell that uses a non‑existent font
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Text using a missing font");
        Style style = workbook.CreateStyle();
        style.Font.Name = "NonExistentFont";
        sheet.Cells["A1"].SetStyle(style);

        // Configure PDF save options and attach the warning callback
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            DefaultFont = "Arial",               // fallback default font
            WarningCallback = new RenderingWarningCallback()
        };

        // Save the workbook to PDF; warnings will be intercepted by the callback
        workbook.Save("ConvertedDocument.pdf", pdfOptions);

        // Output total number of font substitution warnings captured
        var callback = (RenderingWarningCallback)pdfOptions.WarningCallback;
        Console.WriteLine($"Total font substitution warnings: {callback.WarningCount}");
    }
}
using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Configure load options to auto‑fit only rows whose height is not customed
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.AutoFitterOptions = new AutoFitterOptions();
        loadOptions.AutoFitterOptions.OnlyAuto = true;

        // Load the workbook with the above options (replace with your source file)
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Prepare PDF save options (optional: fit entire sheet on one page)
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            OnePagePerSheet = true,
            AllColumnsInOnePagePerSheet = true
        };

        // Save the workbook as PDF; rows are pre‑fitted according to OnlyAuto setting
        workbook.Save("output.pdf", pdfOptions);
    }
}
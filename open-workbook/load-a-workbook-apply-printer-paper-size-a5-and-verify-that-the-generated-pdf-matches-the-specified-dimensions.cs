using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Load an existing workbook and set the default printer paper size to A5 via LoadOptions
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.SetPaperSize(PaperSizeType.PaperA5);
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Ensure the first worksheet's page setup also uses A5
        workbook.Worksheets[0].PageSetup.PaperSize = PaperSizeType.PaperA5;

        // Save the workbook as PDF
        string pdfPath = "output.pdf";
        workbook.Save(pdfPath, SaveFormat.Pdf);

        // Render the workbook to obtain the page size in inches
        ImageOrPrintOptions renderOptions = new ImageOrPrintOptions();
        WorkbookRender renderer = new WorkbookRender(workbook, renderOptions);
        float[] pageSizeInches = renderer.GetPageSizeInch(0); // [0]=width, [1]=height

        Console.WriteLine($"PDF page size: {pageSizeInches[0]:F2} x {pageSizeInches[1]:F2} inches");

        // Expected A5 dimensions (148 mm x 210 mm) converted to inches
        const float expectedWidthInches = 5.83f;  // 148 mm ≈ 5.83 in
        const float expectedHeightInches = 8.27f; // 210 mm ≈ 8.27 in

        bool matchesA5 = Math.Abs(pageSizeInches[0] - expectedWidthInches) < 0.1f &&
                        Math.Abs(pageSizeInches[1] - expectedHeightInches) < 0.1f;

        Console.WriteLine($"Dimensions match A5: {matchesA5}");
    }
}
using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Custom LoadFilter that loads only the specified sheet indexes
class CustomLoadFilter : LoadFilter
{
    private readonly int[] _sheetsOrder;

    public CustomLoadFilter(int[] sheetsOrder) : base(LoadDataFilterOptions.All)
    {
        _sheetsOrder = sheetsOrder;
    }

    // Return the desired sheet indexes; sheets not in this array will be skipped
    public override int[] SheetsInLoadingOrder => _sheetsOrder;
}

class LoadSpecificSheetsAndExportPdf
{
    static void Main()
    {
        // Path to the source workbook
        string sourceFile = "InputWorkbook.xlsx";

        // Define which sheet indexes to load (0‑based)
        int[] sheetsToLoad = new int[] { 0, 2 };

        // Create LoadOptions and assign the custom filter
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFilter = new CustomLoadFilter(sheetsToLoad);

        // Load the workbook with the filter applied
        Workbook workbook = new Workbook(sourceFile, loadOptions);

        // Export each loaded sheet to a separate PDF file
        foreach (int sheetIndex in sheetsToLoad)
        {
            // Ensure the sheet index exists in the loaded workbook
            if (sheetIndex < workbook.Worksheets.Count)
            {
                // Configure PDF save options to include only the current sheet
                PdfSaveOptions pdfOptions = new PdfSaveOptions();
                pdfOptions.SheetSet = new SheetSet(new int[] { sheetIndex });

                // Build output file name (e.g., Sheet1.pdf, Sheet3.pdf)
                string outputFile = $"Sheet{sheetIndex + 1}.pdf";

                // Save the selected sheet as PDF
                workbook.Save(outputFile, pdfOptions);
                Console.WriteLine($"Saved sheet index {sheetIndex} to {outputFile}");
            }
            else
            {
                Console.WriteLine($"Sheet index {sheetIndex} does not exist in the workbook.");
            }
        }
    }
}
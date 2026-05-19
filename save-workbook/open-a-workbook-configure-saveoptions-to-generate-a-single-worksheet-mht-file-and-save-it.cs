using System;
using Aspose.Cells;

class GenerateMhtSingleWorksheet
{
    static void Main()
    {
        // Path to the source Excel file
        string inputPath = "input.xlsx";

        // Path for the resulting MHT file (single worksheet)
        string outputPath = "single_sheet.mht";

        // Load the workbook from the existing file
        Workbook workbook = new Workbook(inputPath);

        // Ensure the first worksheet is the active one (optional)
        workbook.Worksheets.ActiveSheetIndex = 0;

        // Create HtmlSaveOptions for MHTML format
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.MHtml);

        // Export only the active worksheet to the MHT file
        saveOptions.ExportActiveWorksheetOnly = true;

        // Save as a single file (MHTML is inherently single, but set for completeness)
        saveOptions.SaveAsSingleFile = true;

        // Save the workbook using the configured options
        workbook.Save(outputPath, saveOptions);
    }
}
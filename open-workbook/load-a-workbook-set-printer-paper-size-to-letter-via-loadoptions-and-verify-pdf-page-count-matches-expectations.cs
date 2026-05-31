using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Paths to the source Excel file and the output PDF file
        string inputPath = "input.xlsx";
        string outputPdf = "output.pdf";

        // Create LoadOptions and set the default printer paper size to Letter
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.SetPaperSize(PaperSizeType.PaperLetter);

        // Load the workbook using the specified LoadOptions
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Create ImageOrPrintOptions (default settings are sufficient for preview)
        ImageOrPrintOptions printOptions = new ImageOrPrintOptions();

        // Evaluate the total page count of the workbook based on the current print settings
        WorkbookPrintingPreview preview = new WorkbookPrintingPreview(workbook, printOptions);
        int evaluatedPageCount = preview.EvaluatedPageCount;

        // Define the expected page count (adjust as needed for your scenario)
        int expectedPageCount = 1;

        // Verify that the evaluated page count matches the expectation
        Console.WriteLine($"Evaluated page count: {evaluatedPageCount}");
        Console.WriteLine($"Expected page count: {expectedPageCount}");
        Console.WriteLine(evaluatedPageCount == expectedPageCount
            ? "Page count matches expectation."
            : "Page count does NOT match expectation.");

        // Save the workbook as a PDF file
        workbook.Save(outputPdf, SaveFormat.Pdf);
    }
}
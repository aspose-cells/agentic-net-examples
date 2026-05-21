using System;
using Aspose.Cells;

class LoadSetPaperSizeAndSavePdf
{
    static void Main()
    {
        // Path to the source workbook (replace with your actual file)
        string sourceFile = "input.xlsx";

        // Create LoadOptions and set the default print paper size to A3
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.SetPaperSize(PaperSizeType.PaperA3);

        // Load the workbook using the specified LoadOptions (LightCells API is engaged via LoadOptions)
        Workbook workbook = new Workbook(sourceFile, loadOptions);

        // Save the workbook as PDF; the paper size setting will be applied during rendering
        workbook.Save("output.pdf", SaveFormat.Pdf);

        Console.WriteLine("Workbook loaded, paper size set to A3, and saved as PDF.");
    }
}
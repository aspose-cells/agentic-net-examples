using System;
using Aspose.Cells;

class HtmlTableToMergedCells
{
    static void Main()
    {
        // Path to the HTML file that contains tables with colspan and rowspan
        string htmlPath = "input.html";

        // Load the HTML file into a workbook.
        // Aspose.Cells automatically converts HTML tables to worksheets
        // and merges cells according to colspan/rowspan attributes.
        HtmlLoadOptions loadOptions = new HtmlLoadOptions();
        Workbook workbook = new Workbook(htmlPath, loadOptions);

        // Save the resulting workbook to an Excel file.
        string outputPath = "output.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);

        Console.WriteLine("HTML tables have been converted to merged cells and saved to " + outputPath);
    }
}
using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source Excel file (XLSX)
        string sourcePath = "input.xlsx";

        // Path for the output HTML file
        string outputPath = "output.html";

        // Load the workbook from the XLSX file
        Workbook workbook = new Workbook(sourcePath);

        // Create HTML save options and enable presentation preference for better layout
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.PresentationPreference = true;

        // Save the workbook as HTML using the configured options
        workbook.Save(outputPath, htmlOptions);

        Console.WriteLine("Conversion completed. HTML saved to: " + outputPath);
    }
}
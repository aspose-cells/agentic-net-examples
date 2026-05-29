using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add some sample data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello, Aspose.Cells!");

        // Configure HTML save options to disable downlevel-revealed conditional comments
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.DisableDownlevelRevealedComments = true;

        // Save the workbook as an HTML file using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}
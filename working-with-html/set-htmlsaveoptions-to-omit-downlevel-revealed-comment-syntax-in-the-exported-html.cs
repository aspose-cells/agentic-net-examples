using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample data (optional)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample data");

            // Configure HTML save options to omit downlevel revealed comments
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.DisableDownlevelRevealedComments = true;

            // Save the workbook as HTML using the configured options
            workbook.Save("output.html", htmlOptions);

            Console.WriteLine("Workbook saved with DisableDownlevelRevealedComments = true");
        }
    }
}
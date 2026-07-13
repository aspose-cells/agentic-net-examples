using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create an empty workbook (or load an existing one as needed)
        Workbook workbook = new Workbook();

        // Configure HtmlSaveOptions
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // NOTE: Aspose.Cells currently does not provide a dedicated HtmlSaveOptions property
        // to collapse multiple spaces after newline characters during HTML export.
        // If such a feature exists in a newer version, replace the placeholder below
        // with the appropriate property assignment.
        // Example (hypothetical): htmlOptions.CollapseSpacesAfterNewline = true;

        // Save the workbook as HTML using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}
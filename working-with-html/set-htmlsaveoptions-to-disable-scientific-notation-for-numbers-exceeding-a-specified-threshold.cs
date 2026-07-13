using System;
using Aspose.Cells;
using Aspose.Cells.Saving;

class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Create HtmlSaveOptions
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

        // The Aspose.Cells API provides properties to control scientific notation
        // in HTML output, but these properties are not present in the supplied
        // documentation. The intended usage is shown below as a placeholder.
        // Uncomment and adjust when the correct members are confirmed.

        // htmlOptions.DisableScientificNotation = true;               // Disable scientific notation
        // htmlOptions.ScientificNotationThreshold = 1e12;            // Threshold for disabling

        // Save the workbook as HTML using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}

// Author: Aspose.Cells .NET example – placeholder for scientific notation settings (pending API verification)
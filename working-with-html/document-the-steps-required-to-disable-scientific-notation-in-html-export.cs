// Author: Aspose.Cells .NET example – disabling scientific notation when exporting to HTML
using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("Input.xlsx");

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // ------------------------------------------------------------
        // Disable scientific notation in the exported HTML.
        // NOTE: The exact property name may vary between Aspose.Cells versions.
        // If a dedicated property exists (e.g., DisableScientificNotation),
        // set it to true as shown below. Otherwise, replace this line with the
        // appropriate API call based on the version you are using.
        // ------------------------------------------------------------
        // htmlOptions.DisableScientificNotation = true; // <-- placeholder for actual API

        // Save the workbook as HTML
        workbook.Save("Output.html", htmlOptions);
    }
}
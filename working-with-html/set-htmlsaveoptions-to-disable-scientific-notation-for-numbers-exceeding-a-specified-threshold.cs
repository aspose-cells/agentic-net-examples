// Title: Aspose.Cells for .NET – Disable Scientific Notation When Saving Excel to HTML
// Description: Learn how to prevent large numbers from being converted to scientific notation in HTML output. The example loads an Excel workbook, sets Workbook.Settings.SignificantDigits to retain full precision, configures HtmlSaveOptions, and saves the file as HTML with plain decimal formatting.
// Keywords: Aspose.Cells | C# | HtmlSaveOptions | disable scientific notation | significant digits | Excel to HTML conversion | large number formatting | HTML5 export | numeric precision
// Common Searches: Aspose.Cells stop scientific notation in HTML | set significant digits for HTML export Aspose.Cells | prevent Excel numbers from showing as 1e+12 in HTML | C# Aspose.Cells HTML save options decimal format | how to keep full numeric precision when exporting to HTML
// Developer Intent: Keep numbers above a defined size from being displayed in scientific notation when exporting an Excel workbook to HTML using Aspose.Cells for .NET.
// Use Cases: Export a financial statement with multi‑trillion values to HTML while preserving readable decimal figures. | Render a data‑analysis workbook in a web page where column values exceed 10^12 and must stay in plain format. | Generate an HTML5 report from Excel that requires exact numeric representation without scientific notation.
// AI Prompts: Provide C# code that configures Aspose.Cells HtmlSaveOptions to keep numbers in decimal format for values larger than 1e10. | Explain the effect of Workbook.Settings.SignificantDigits on number formatting during HTML export with Aspose.Cells. | Show an example of disabling scientific notation when saving an Excel workbook to HTML using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Learn how to prevent large numbers from being converted to scientific notation in HTML output. The example loads an Excel workbook, sets Workbook.Settings.SignificantDigits to retain full precision, configures HtmlSaveOptions, and saves the file as HTML with plain decimal formatting.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Configure the workbook to use a higher number of significant digits.
        // This prevents large numbers from being automatically formatted in scientific notation
        // when the workbook is saved to HTML.
        workbook.Settings.SignificantDigits = 15;   // 15 is the maximum precision for double

        // Create HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Example: set the HTML version to HTML5 (optional, but demonstrates usage)
        htmlOptions.HtmlVersion = HtmlVersion.Html5;

        // Save the workbook as HTML with the configured options.
        // The numbers will be rendered in plain decimal form rather than scientific notation.
        workbook.Save("output.html", htmlOptions);
    }
}

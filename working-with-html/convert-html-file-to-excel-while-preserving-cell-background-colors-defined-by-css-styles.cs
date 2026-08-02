// Title: HTML to Excel conversion with CSS background colors – Aspose.Cells for .NET
// Description: Shows how to load an HTML file that contains CSS styling (including cell background colors) into an Aspose.Cells Workbook using LoadOptions.LoadFormat.Html, then save it as an XLSX workbook with C#.
// Keywords: Aspose.Cells HTML to Excel | C# preserve CSS background color | LoadFormat.Html | HTML table to XLSX | convert styled HTML to Excel | Aspose.Cells .NET | cell background color conversion | HTML to XLSX C#
// Common Searches: Aspose.Cells keep cell colors when converting HTML to Excel | C# load HTML with CSS into workbook | HTML table to Excel preserving formatting Aspose | Export styled HTML as XLSX using Aspose.Cells | LoadOptions LoadFormat.Html background color
// Developer Intent: Convert an HTML document into an Excel workbook while retaining the CSS‑defined cell background colors.
// Use Cases: Create Excel reports from web‑based tables that maintain the original color scheme for corporate dashboards. | Automate batch conversion of styled HTML files to XLSX files for offline analysis without losing visual cues. | Integrate HTML‑to‑Excel export in a .NET application where the generated spreadsheets must reflect the source page’s formatting.
// AI Prompts: Generate C# code with Aspose.Cells that loads an HTML file and saves it as XLSX, ensuring CSS background colors are applied to cells. | Explain the role of LoadOptions.LoadFormat.Html in preserving CSS formatting, especially background colors, when loading HTML into a Workbook. | Provide a method to verify and, if needed, adjust cell background colors after importing an HTML file with Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to load an HTML file that contains CSS styling (including cell background colors) into an Aspose.Cells Workbook using LoadOptions.LoadFormat.Html, then save it as an XLSX workbook with C#.
class HtmlToExcelConverter
{
    static void Main()
    {
        // Path to the source HTML file
        string htmlFilePath = "input.html";

        // Path for the resulting Excel file
        string excelFilePath = "output.xlsx";

        // Load the HTML file into a Workbook.
        // LoadOptions with LoadFormat.Html ensures the HTML is parsed correctly,
        // including CSS styles such as cell background colors.
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
        Workbook workbook = new Workbook(htmlFilePath, loadOptions);

        // Save the workbook as an Excel file (XLSX format).
        workbook.Save(excelFilePath, SaveFormat.Xlsx);

        Console.WriteLine("Conversion completed. Excel file saved to: " + excelFilePath);
    }
}

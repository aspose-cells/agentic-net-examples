// Title: C# – Load Excel, Modify Cell B4, Export to HTML with a Custom Default Font (Aspose.Cells)
// Description: Using Aspose.Cells for .NET, this example loads an Excel workbook, updates cell B4, sets HtmlSaveOptions.DefaultFontName to a custom font, and saves the result as an HTML file.
// Keywords: Aspose.Cells | C# Excel to HTML | HtmlSaveOptions | DefaultFontName | modify cell value | export Excel as HTML | custom default font | load workbook | save as HTML | Aspose.Cells example
// Common Searches: Aspose.Cells change cell value and save as HTML | C# set default font for HTML export using Aspose.Cells | How to use HtmlSaveOptions.DefaultFontName | Convert Excel to HTML with custom font C# | Update cell B4 programmatically Aspose.Cells
// Developer Intent: Change the content of cell B4 and generate an HTML file that uses a specified default font.
// Use Cases: Create web‑compatible spreadsheet previews that match site typography. | Automate report generation where specific cells are updated before HTML conversion. | Produce consistent HTML output for documentation portals by enforcing a default font. | Integrate Excel‑to‑HTML conversion into a .NET web application with custom styling.
// AI Prompts: Write C# code with Aspose.Cells to change cell C5, set the HTML default font to Arial, and save the workbook as HTML. | Explain how HtmlSaveOptions.DefaultFontName affects the generated HTML and how to combine it with external CSS for further styling. | Provide a step‑by‑step guide to load an Excel file, modify multiple cells, and export to HTML with a custom default font using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Using Aspose.Cells for .NET, this example loads an Excel workbook, updates cell B4, sets HtmlSaveOptions.DefaultFontName to a custom font, and saves the result as an HTML file.
class Program
{
    static void Main()
    {
        // Load the existing Excel file
        // (Assumes "input.xlsx" exists in the same directory as the executable)
        Workbook workbook = new Workbook("input.xlsx");

        // Modify the text of cell B4 in the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["B4"].PutValue("Modified text for B4");

        // Create HTML save options and set a custom default font
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.DefaultFontName = "Courier New"; // Custom default font

        // Save the workbook as an HTML file using the specified options
        workbook.Save("output.html", htmlOptions);

        Console.WriteLine("Excel file has been converted to HTML with custom default font.");
    }
}

// Title: Load an Excel workbook, change cell B4 text, and export to HTML with a custom default font using Aspose.Cells for .NET
// AI Prompts: Load input.xlsx with Aspose.Cells, set worksheet.Cells["B4"] to "New Text", configure HtmlSaveOptions.DefaultFont to "Arial", and save as output.html. | Read an existing .xlsx file, update a specific cell value, apply a default Arial font via HtmlSaveOptions, and generate an HTML file using C#.
// Common Searches: Aspose.Cells C# change cell value and save workbook as HTML with specific font | how to set default font in HtmlSaveOptions when converting Excel to HTML | update B4 cell in Excel and export to HTML using Aspose.Cells .NET | set Arial as default font for HTML output from Aspose.Cells workbook | convert xlsx to html with custom font using Aspose.Cells C# example
// Tags: Aspose.Cells HtmlSaveOptions default font | modify worksheet cell before HTML export | Excel to HTML conversion with custom font | C# update cell B4 Aspose.Cells | set default font for HTML output Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;
using System.Reflection;

// The program checks for input.xlsx, loads it with Aspose.Cells, changes cell B4 to "New Text", uses reflection to set HtmlSaveOptions.DefaultFont to Arial for compatibility across library versions, and saves the workbook as output.html.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing Excel file
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Modify the text of cell B4
            Cell cell = worksheet.Cells["B4"];
            cell.PutValue("New Text");

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Set default font if the property exists (covers different library versions)
            PropertyInfo defaultFontProp = typeof(HtmlSaveOptions).GetProperty("DefaultFont");
            if (defaultFontProp != null && defaultFontProp.CanWrite)
            {
                defaultFontProp.SetValue(htmlOptions, "Arial");
            }

            // Save the workbook as an HTML file using the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

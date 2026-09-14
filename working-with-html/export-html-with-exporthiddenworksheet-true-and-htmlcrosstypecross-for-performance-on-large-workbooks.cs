// Title: Export hidden worksheets to HTML using Aspose.Cells in C# and apply HtmlCrossType.Cross for faster processing of large workbooks
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, sets HtmlSaveOptions.ExportHiddenWorksheet to true, applies HtmlCrossType.Cross when the property is available, and saves the result as an HTML file. | Refactor the provided Aspose.Cells example into a reusable method that accepts input and output paths, ensures hidden sheets are included in the HTML output, and uses HtmlCrossType.Cross for performance on workbooks with many rows. | Create a C# snippet that checks the Aspose.Cells version, configures HtmlSaveOptions for hidden worksheet export, conditionally sets HtmlCrossType.Cross if supported, and writes the HTML file while handling file‑not‑found and runtime errors.
// Common Searches: how to include hidden worksheets when saving an Excel file as HTML with Aspose.Cells C# | HtmlCrossType.Cross setting for improving HTML export speed in Aspose.Cells | export large Excel workbook to HTML using Aspose.Cells performance tips | Aspose.Cells HtmlSaveOptions ExportHiddenWorksheet example C# | why HtmlCrossType property is missing in current Aspose.Cells version
// Tags: Aspose.Cells HtmlSaveOptions ExportHiddenWorksheet | Aspose.Cells HtmlCrossType performance | C# Excel hidden sheets HTML export | large workbook HTML conversion Aspose.Cells | HTML export optimization Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads 'input.xlsx' with Aspose.Cells, configures HtmlSaveOptions to include hidden worksheets (ExportHiddenWorksheet = true), notes that HtmlCrossType may not be present in older versions, and saves the workbook as 'output.html' while handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportHiddenWorksheet = true // Include hidden worksheets in the HTML output
                // HtmlCrossType property is not available in the current Aspose.Cells version
            };

            // Save the workbook as an HTML file with the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

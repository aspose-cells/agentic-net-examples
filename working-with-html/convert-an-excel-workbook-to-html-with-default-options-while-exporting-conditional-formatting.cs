// Title: Convert Excel to HTML with default Aspose.Cells settings and preserve conditional formatting (C#)
// Description: A C# sample that loads an .xlsx workbook with Aspose.Cells, applies the default HtmlSaveOptions (explicitly setting MergeAreas=true to keep conditional‑formatting ranges merged), and uses ConversionUtility.Convert to produce an HTML file.
// Keywords: Aspose.Cells | C# | Excel to HTML conversion | HtmlSaveOptions | MergeAreas | conditional formatting export | ConversionUtility | LoadOptions | Xlsx to HTML | default HTML save options
// Common Searches: Aspose.Cells export Excel to HTML with conditional formatting | C# convert .xlsx to HTML preserving colors | HtmlSaveOptions default settings example | How to use ConversionUtility.Convert for HTML output | MergeAreas true Aspose.Cells HTML conversion
// Developer Intent: Create an HTML representation of an Excel workbook while retaining its conditional‑formatting using Aspose.Cells default options.
// Use Cases: Display financial dashboards on a web page with the same color‑coded rules as the original Excel file. | Batch‑process uploaded spreadsheets to generate previewable HTML reports in a portal, ensuring formatting consistency. | Produce documentation from Excel templates where default styling and conditional highlights must appear in the HTML output.
// AI Prompts: Generate C# code that converts an .xlsx file to HTML with Aspose.Cells, keeping conditional formatting and using default HtmlSaveOptions. | Show how to enable HtmlSaveOptions.MergeAreas for preserving conditional‑formatting ranges during Excel‑to‑HTML conversion. | Explain the step‑by‑step process of loading a workbook, configuring HtmlSaveOptions, and calling ConversionUtility.Convert to export HTML in .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// A C# sample that loads an .xlsx workbook with Aspose.Cells, applies the default HtmlSaveOptions (explicitly setting MergeAreas=true to keep conditional‑formatting ranges merged), and uses ConversionUtility.Convert to produce an HTML file.
class ExcelToHtmlConverter
{
    static void Main()
    {
        // Path to the source Excel file
        string sourcePath = "input.xlsx";

        // Path where the HTML output will be saved
        string htmlPath = "output.html";

        // Load options for the source workbook (auto-detect format if needed)
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

        // Create HTML save options with default settings
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Ensure conditional formatting areas are merged (default is true, set explicitly for clarity)
        htmlOptions.MergeAreas = true;

        // Convert the Excel file to HTML using the conversion utility
        ConversionUtility.Convert(sourcePath, loadOptions, htmlPath, htmlOptions);

        Console.WriteLine("Conversion completed. HTML saved to: " + htmlPath);
    }
}

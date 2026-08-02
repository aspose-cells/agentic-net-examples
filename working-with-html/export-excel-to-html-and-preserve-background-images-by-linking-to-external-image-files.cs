// Title: Export Excel to HTML with external background images using Aspose.Cells for .NET (C#)
// Description: Loads an .xlsx workbook, configures HtmlSaveOptions to turn off Base64 image embedding, defines an attached files folder and URL prefix, enables automatic folder creation, and saves the workbook as HTML so worksheet background images are linked as separate files.
// Keywords: Aspose.Cells | C# | .NET | Excel to HTML | ExportImagesAsBase64 false | HtmlSaveOptions | external image files | background image export | attached files directory | HTML conversion with linked images
// Common Searches: Aspose.Cells export Excel to HTML external images | How to save Excel background images as separate files with Aspose.Cells | HtmlSaveOptions ExportImagesAsBase64 false C# example | Set attached files directory for HTML export in Aspose.Cells | C# convert workbook to HTML with linked background graphics | Aspose.Cells HTML export folder creation option
// Developer Intent: Generate an HTML version of an Excel workbook where all background and embedded images are written to disk as separate files and referenced via relative URLs.
// Use Cases: Publish Excel‑based reports on the web while keeping background graphics cacheable as independent image files. | Integrate Excel‑to‑HTML conversion into a web API that serves HTML pages with images stored in a dedicated assets folder. | Batch‑process multiple workbooks, preserving their visual assets in a structured output directory for downstream publishing pipelines. | Create lightweight HTML snapshots of spreadsheets for documentation portals where image size and load time matter.
// AI Prompts: Show how to customize the naming pattern of exported image files when using HtmlSaveOptions in Aspose.Cells. | Provide a code snippet that sets a custom URL prefix for attached images and disables automatic directory creation. | Explain strategies for handling existing image files in the output folder when re‑exporting the same workbook to HTML.

using System;
using Aspose.Cells;

// Loads an .xlsx workbook, configures HtmlSaveOptions to turn off Base64 image embedding, defines an attached files folder and URL prefix, enables automatic folder creation, and saves the workbook as HTML so worksheet background images are linked as separate files.
class ExportExcelToHtml
{
    static void Main()
    {
        // Load the Excel workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();

        // Export images as separate files (not Base64)
        saveOptions.ExportImagesAsBase64 = false;

        // Folder where the image files will be written
        saveOptions.AttachedFilesDirectory = "output_files";

        // URL prefix used in the generated HTML to reference the images
        saveOptions.AttachedFilesUrlPrefix = "output_files/";

        // Automatically create the folder if it does not exist
        saveOptions.CreateDirectory = true;

        // Save the workbook as HTML; background images will be linked to external files
        workbook.Save("output.html", saveOptions);
    }
}

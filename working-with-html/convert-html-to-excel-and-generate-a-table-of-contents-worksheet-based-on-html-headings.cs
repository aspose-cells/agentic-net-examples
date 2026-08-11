// Title: C# – Convert HTML to Excel with Aspose.Cells and Add a Table of Contents Sheet
// Description: Loads an HTML file into an Aspose.Cells Workbook, creates a "Table of Contents" worksheet, lists each generated sheet with internal hyperlinks to its A1 cell, and saves the result as an XLSX file.
// Keywords: Aspose.Cells HTML to Excel C# | generate TOC worksheet Aspose | internal hyperlink Excel sheet | load HTML workbook Aspose.Cells | programmatic Excel Table of Contents
// Common Searches: Aspose.Cells convert HTML to XLSX C# | create table of contents sheet in Excel using Aspose | add internal hyperlinks between worksheets Aspose.Cells | C# code to load HTML into workbook and generate TOC | Aspose.Cells example for HTML headings to Excel
// Developer Intent: Produce an Excel file from an HTML document and automatically add a navigable Table of Contents worksheet.
// Use Cases: Transform a web‑based report into an Excel workbook with quick navigation links. | Export segmented documentation (one HTML section per sheet) and provide a clickable index. | Package data extracted from a website into a downloadable XLSX file that includes a TOC for end‑users.
// AI Prompts: Show how to indent TOC entries based on HTML heading levels (H1, H2, H3). | Provide code to apply custom styling (font color, underline) to the TOC hyperlinks. | Explain how to add row numbers or page references next to each TOC entry.

using System;
using System.IO;
using Aspose.Cells;

// Loads an HTML file into an Aspose.Cells Workbook, creates a "Table of Contents" worksheet, lists each generated sheet with internal hyperlinks to its A1 cell, and saves the result as an XLSX file.
class HtmlToExcelWithToc
{
    static void Main()
    {
        try
        {
            // Path to the source HTML file
            string htmlPath = "input.html";

            // Verify that the HTML file exists to avoid FileNotFoundException
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Error: The file \"{htmlPath}\" was not found.");
                return;
            }

            // Load the HTML file into a workbook using HtmlLoadOptions
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Add a new worksheet that will serve as the Table of Contents
            int tocIndex = workbook.Worksheets.Add();
            Worksheet tocSheet = workbook.Worksheets[tocIndex];
            tocSheet.Name = "Table of Contents";

            // Populate the TOC with links to each worksheet (excluding the TOC itself)
            int tocRow = 0;
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet ws = workbook.Worksheets[i];
                if (ws.Name == "Table of Contents")
                    continue;

                // Write the sheet name in the TOC
                tocSheet.Cells[tocRow, 0].PutValue(ws.Name);

                // Add an internal hyperlink that points to cell A1 of the target sheet
                // Correct overload: Add(firstRow, firstColumn, totalRows, totalColumns, address)
                tocSheet.Hyperlinks.Add(tocRow, 0, 1, 1, ws.Name + "!A1");

                tocRow++;
            }

            // Save the workbook as an Excel file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

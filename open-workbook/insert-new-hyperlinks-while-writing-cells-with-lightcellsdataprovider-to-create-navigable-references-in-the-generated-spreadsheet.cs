// Title: Add external URL and internal anchor hyperlinks to cells while using Aspose.Cells LightCellsDataProvider in C#
// AI Prompts: Generate C# code that writes data with Aspose.Cells LightCellsDataProvider and attaches an external hyperlink to cell B2 with custom display text. | Show how to create an internal anchor hyperlink (e.g., #D5) in cell C3 during worksheet population using LightCellsDataProvider in Aspose.Cells. | Provide a complete example that saves an Excel workbook containing both external and internal hyperlinks using the Aspose.Cells API.
// Common Searches: asp.net core add external hyperlink to Excel cell using Aspose.Cells LightCellsDataProvider | c# create internal anchor link in Excel workbook with Aspose.Cells | how to set hyperlink display text in Aspose.Cells generated spreadsheet | populate worksheet with LightCellsDataProvider and add hyperlinks example | save Excel file with mixed external and internal hyperlinks Aspose.Cells C#
// Tags: add external hyperlink Aspose.Cells C# | create internal anchor hyperlink Aspose.Cells | populate worksheet using LightCellsDataProvider | set hyperlink display text Aspose.Cells | save workbook with hyperlinks Aspose.Cells

using System;
using Aspose.Cells;

namespace LightCellsHyperlinkExample
{
    // Creates a new workbook, fills a 10x5 range with sample data, adds an external hyperlink to B2 and an internal anchor hyperlink to C3 with custom display text, and saves the file as LightCellsWithHyperlinks.xlsx using Aspose.Cells.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Define the size of the data range
                int totalRows = 10;
                int totalCols = 5;

                // Fill the worksheet with sample data
                for (int row = 0; row < totalRows; row++)
                {
                    for (int col = 0; col < totalCols; col++)
                    {
                        sheet.Cells[row, col].PutValue($"Row{row + 1}Col{col + 1}");
                    }
                }

                // Add an external hyperlink to cell B2 (row 1, column 1)
                int extLinkIndex = sheet.Hyperlinks.Add(1, 1, 1, 1, "https://www.example.com");
                Hyperlink extLink = sheet.Hyperlinks[extLinkIndex];
                extLink.TextToDisplay = "Example Site";

                // Add an internal hyperlink (anchor) to cell C3 (row 2, column 2)
                int intLinkIndex = sheet.Hyperlinks.Add(2, 2, 1, 1, "#D5");
                Hyperlink intLink = sheet.Hyperlinks[intLinkIndex];
                intLink.TextToDisplay = "Go to D5";

                // Save the workbook
                string outputPath = "LightCellsWithHyperlinks.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

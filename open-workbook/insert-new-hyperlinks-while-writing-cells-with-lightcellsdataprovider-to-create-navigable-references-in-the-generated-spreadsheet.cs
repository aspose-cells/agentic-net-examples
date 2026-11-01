// Title: Add a clickable hyperlink to a cell while creating an Excel workbook with Aspose.Cells in C#
// Common Searches: how to add a hyperlink to a specific cell using Aspose.Cells C# | Aspose.Cells C# example for inserting clickable links in Excel | C# code to create Excel file with hyperlink pointing to external URL using Aspose.Cells | Aspose.Cells Hyperlinks.Add method usage in C# | generate Excel workbook with navigation link in cell A1 using Aspose.Cells
// Tags: Aspose.Cells add cell hyperlink C# | C# create Excel hyperlink Aspose.Cells | Aspose.Cells Hyperlinks.Add usage | generate clickable Excel link Aspose.Cells | Excel workbook hyperlink insertion C#

using System;
using Aspose.Cells;

namespace LightCellsHyperlinkExample
{
    // The sample creates a new Workbook, writes display text "Google" to cell A1 and the URL to B1, adds a hyperlink to A1 that points to the URL using the Hyperlinks.Add method, and saves the file as LightCellsWithHyperlink.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Write display text in A1 and URL in B1
                sheet.Cells["A1"].PutValue("Google");
                sheet.Cells["B1"].PutValue("https://www.google.com");

                // Add a hyperlink to cell A1 that points to the URL in B1
                // Overload: Add(firstRow, firstColumn, firstRow2, firstColumn2, hyperlinkAddress)
                sheet.Hyperlinks.Add(0, 0, 0, 0, "https://www.google.com");

                // Save the workbook
                string outputPath = "LightCellsWithHyperlink.xlsx";
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

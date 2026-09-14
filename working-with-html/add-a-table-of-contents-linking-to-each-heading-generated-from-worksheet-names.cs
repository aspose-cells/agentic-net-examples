// Title: How to generate a Table of Contents worksheet with clickable links to each sheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to insert a 'Table of Contents' sheet at the start of a workbook and adds hyperlinks from each entry to the A1 cell of the corresponding worksheet. | Show an example that creates several worksheets, builds a TOC sheet, writes sheet names with a bold header, applies auto‑fit to the column, and saves the file as an .xlsx. | Demonstrate looping through all worksheets (excluding the TOC) and adding Hyperlink objects that point to each sheet's first cell using the Aspose.Cells API.
// Common Searches: aspnet c# add table of contents to Excel workbook with Aspose.Cells | aspose.cells create worksheet hyperlinks for a TOC sheet | generate Excel TOC sheet programmatically using Aspose.Cells .NET | how to auto fit columns after populating data with Aspose.Cells | C# example linking TOC entries to sheet A1 using Aspose.Cells
// Tags: Aspose.Cells generate TOC worksheet | Aspose.Cells add worksheet hyperlinks | Aspose.Cells auto fit columns | Aspose.Cells hyperlink to sheet A1 | Aspose.Cells build TOC from worksheet collection

using Aspose.Cells;
using System;

// The sample creates a new workbook, adds several worksheets, inserts a 'Table of Contents' sheet at the first position, lists each worksheet name (excluding the TOC) in column A, adds hyperlinks to each sheet's A1 cell, auto‑fits the column width, and saves the workbook as WorkbookWithTOC.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample worksheets (replace or remove as needed)
            workbook.Worksheets.Add("Introduction");
            workbook.Worksheets.Add("Data");
            workbook.Worksheets.Add("Analysis");
            workbook.Worksheets.Add("Conclusion");

            // Insert a Table of Contents sheet at the first position
            Worksheet tocSheet = workbook.Worksheets[0];
            tocSheet.Name = "Table of Contents";

            // Header for TOC
            tocSheet.Cells["A1"].PutValue("Table of Contents");
            Style headerStyle = tocSheet.Cells["A1"].GetStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.Font.Size = 14;
            tocSheet.Cells["A1"].SetStyle(headerStyle);

            // List each worksheet (skip the TOC itself) and add a hyperlink to its A1 cell
            int row = 2; // start from row 2
            for (int i = 1; i < workbook.Worksheets.Count; i++)
            {
                Worksheet ws = workbook.Worksheets[i];

                // Write sheet name
                tocSheet.Cells[row, 0].PutValue(ws.Name);

                // Add hyperlink pointing to the sheet's A1 cell
                // Hyperlinks.Add(row, column, totalRows, totalColumns, address)
                tocSheet.Hyperlinks.Add(row, 0, 1, 1, ws.Name + "!A1");

                row++;
            }

            // Adjust column width to fit content
            tocSheet.AutoFitColumns();

            // Save the workbook
            workbook.Save("WorkbookWithTOC.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}

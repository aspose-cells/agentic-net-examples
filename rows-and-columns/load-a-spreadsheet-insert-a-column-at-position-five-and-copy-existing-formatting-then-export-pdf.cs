// Title: Insert a column at index 5, copy left column formatting, and export the sheet to PDF using Aspose.Cells for .NET
// AI Prompts: Insert a new column at zero‑based index 4 in the first worksheet, duplicate the formatting from column D to the new column, and generate a PDF file with Aspose.Cells. | Add a column after the fourth column, copy only the style attributes from the preceding column, then convert the modified workbook to PDF using the Aspose.Cells ConversionUtility.
// Common Searches: how to insert a column at a specific position and copy its format using Aspose.Cells C# | Aspose.Cells C# copy column formatting without values before PDF conversion | insert column index 4 and preserve styles then export to PDF with Aspose.Cells | C# Aspose.Cells convert modified Excel workbook to PDF after column insertion
// Tags: insert column Aspose.Cells C# | copy column style Aspose.Cells | excel to pdf conversion Aspose.Cells .NET | preserve cell styles Aspose.Cells | column insertion index zero based Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// The example loads an existing Excel file, inserts a new column at zero‑based index 4, copies only the formatting from the adjacent column, saves the updated workbook, and then converts it to a PDF using Aspose.Cells' ConversionUtility.
class Program
{
    static void Main()
    {
        // Paths – adjust as needed
        string inputFile = "input.xlsx";          // existing spreadsheet
        string modifiedFile = "modified.xlsx";    // temporary file after column insertion
        string outputPdf = "output.pdf";          // final PDF file

        // Load the existing workbook
        Workbook workbook = new Workbook(inputFile);

        // Work with the first worksheet (index 0)
        Worksheet sheet = workbook.Worksheets[0];

        // Insert a new column at position five (zero‑based index 4)
        sheet.Cells.InsertColumn(4);

        // Copy formatting from the column to the left (index 3) to the newly inserted column (index 4)
        PasteOptions pasteOptions = new PasteOptions
        {
            PasteType = PasteType.Formats   // copy only formats, not values
        };
        sheet.Cells.CopyColumns(sheet.Cells, 3, 4, 1, pasteOptions);

        // Save the modified workbook to a temporary file
        workbook.Save(modifiedFile);

        // Convert the modified workbook to PDF using the provided conversion utility
        ConversionUtility.Convert(modifiedFile, outputPdf);
    }
}

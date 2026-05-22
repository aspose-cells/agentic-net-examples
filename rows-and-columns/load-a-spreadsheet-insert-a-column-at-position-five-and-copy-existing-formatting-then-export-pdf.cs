using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Paths for the original file, the intermediate modified file, and the final PDF.
        string inputPath = "input.xlsx";
        string intermediatePath = "modified.xlsx";
        string outputPdf = "output.pdf";

        // Load the existing workbook.
        Workbook workbook = new Workbook(inputPath);
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Insert a new column at position five (zero‑based index 4).
        int insertIndex = 4;
        cells.InsertColumn(insertIndex, true);

        // After insertion, the original column that was at insertIndex is now at insertIndex + 1.
        // Copy its formatting to the newly inserted column.
        PasteOptions pasteOptions = new PasteOptions
        {
            PasteType = PasteType.Formats
        };
        cells.CopyColumns(cells, insertIndex + 1, insertIndex, 1, pasteOptions);

        // Save the modified workbook to a temporary file.
        workbook.Save(intermediatePath);

        // Convert the temporary Excel file to PDF.
        ConversionUtility.Convert(intermediatePath, outputPdf);
    }
}
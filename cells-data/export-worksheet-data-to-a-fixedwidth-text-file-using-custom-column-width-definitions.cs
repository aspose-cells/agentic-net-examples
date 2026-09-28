// Title: How to export an Aspose.Cells worksheet to a fixed‑width text file with custom column character widths in C#
// AI Prompts: Generate C# code that reads a range from an Aspose.Cells worksheet, applies specified column character widths, and writes the data to a fixed‑width .txt file. | Show how to configure column widths in an Aspose.Cells workbook and export the selected cells as a padded‑column text file using StreamWriter.
// Common Searches: c# aspocells export selected range to fixed width txt with custom column sizes | how to define column character width in aspocells before text export | aspocells write excel data to padded text file using streamwriter | example of exporting worksheet to fixed‑width text file in aspocells c#
// Tags: Aspose.Cells fixed-width text export | Aspose.Cells custom column character width | Aspose.Cells padded column text generation | Aspose.Cells ExportDataTable to text file

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The sample creates a workbook, fills three columns with data, sets column widths of 20, 5, and 15 characters, exports the defined range to a DataTable, and writes each row to a fixed‑width text file by truncating or padding each cell to its column width. Finally, the workbook is saved as an .xlsx file.
class ExportFixedWidth
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Age");
        sheet.Cells["C1"].PutValue("Country");
        sheet.Cells["A2"].PutValue("John Doe");
        sheet.Cells["B2"].PutValue(29);
        sheet.Cells["C2"].PutValue("United States");
        sheet.Cells["A3"].PutValue("Anna Smith");
        sheet.Cells["B3"].PutValue(34);
        sheet.Cells["C3"].PutValue("Canada");

        // Define custom column widths (in characters)
        double[] colWidths = new double[] { 20, 5, 15 }; // Widths for columns A, B, C

        // Apply the custom widths to the worksheet columns
        for (int i = 0; i < colWidths.Length; i++)
        {
            sheet.Cells.SetColumnWidth(i, colWidths[i]);
        }

        // Export the defined area (rows 0‑2, columns 0‑2) to a DataTable
        int firstRow = 0;
        int firstColumn = 0;
        int totalRows = 3;
        int totalColumns = 3;
        DataTable dt = sheet.Cells.ExportDataTable(firstRow, firstColumn, totalRows, totalColumns, true);

        // Write the data to a fixed‑width text file
        string outputPath = "FixedWidthExport.txt";
        using (StreamWriter writer = new StreamWriter(outputPath))
        {
            foreach (DataRow row in dt.Rows)
            {
                for (int col = 0; col < totalColumns; col++)
                {
                    string cellText = row[col]?.ToString() ?? string.Empty;
                    int width = (int)Math.Ceiling(colWidths[col]); // Convert to integer character width

                    // Truncate if longer than width, otherwise pad with spaces
                    if (cellText.Length > width)
                        cellText = cellText.Substring(0, width);
                    else
                        cellText = cellText.PadRight(width);

                    writer.Write(cellText);
                }
                writer.WriteLine();
            }
        }

        // Save the workbook (optional, demonstrates standard save lifecycle)
        workbook.Save("WorkbookWithData.xlsx");
    }
}

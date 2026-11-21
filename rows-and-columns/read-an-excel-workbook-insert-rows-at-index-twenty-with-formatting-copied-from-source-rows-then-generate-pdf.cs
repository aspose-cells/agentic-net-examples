// Title: Insert a row at position 20 in an Excel workbook, copy its formatting from another row, and export to PDF using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# how to insert a row at position 20 and retain header style | Insert rows without default formatting and copy style from another row using Aspose.Cells | Export Excel to PDF after adding rows with custom formatting in .NET | Copy row style when inserting rows with Aspose.Cells library in C# | Insert a row at row 20 using zero‑based index in Aspose.Cells example | C# Aspose.Cells insert rows and preserve existing row formatting before PDF conversion
// Tags: insert rows with style copy Aspose.Cells | copy row formatting Aspose.Cells C# | export workbook to PDF Aspose.Cells | zero based row index insertion Aspose.Cells | apply source row style to inserted row Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample creates or loads an Excel file, inserts a new row at the zero‑based index 19 (row 20) without inheriting default styles, copies the formatting from a source row to the inserted row, and then saves the workbook as a PDF, handling missing file and generic errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Ensure the input file exists; create a simple workbook if it does not.
            if (!File.Exists(inputPath))
            {
                var tempWb = new Workbook();
                var tempSheet = tempWb.Worksheets[0];
                tempSheet.Name = "Sheet1";

                // Add a header row for style copying demonstration.
                tempSheet.Cells["A1"].PutValue("Header");
                var headerStyle = tempSheet.Cells["A1"].GetStyle();
                headerStyle.Font.IsBold = true;
                tempSheet.Cells["A1"].SetStyle(headerStyle);

                tempWb.Save(inputPath);
            }

            // Load the existing Excel workbook (ensure the file exists before loading)
            var workbook = new Workbook(inputPath);

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];

            // Zero‑based index where new rows will be inserted (row 20 => index 19)
            int insertIndex = 19;
            int rowsToInsert = 1;

            // Insert rows without automatically copying styles
            sheet.Cells.InsertRows(insertIndex, rowsToInsert, false);

            // Source row whose formatting will be copied (example: first row, index 0)
            var sourceRow = sheet.Cells.Rows[0];

            // Row that was just inserted
            var insertedRow = sheet.Cells.Rows[insertIndex];

            // Copy the formatting (style) from the source row to the inserted row
            // Row style is accessed via GetStyle/SetStyle methods.
            Style sourceStyle = sourceRow.GetStyle();
            insertedRow.SetStyle(sourceStyle);

            // Save the modified workbook as a PDF document
            workbook.Save(outputPath, SaveFormat.Pdf);
        }
        catch (FileNotFoundException fnfEx)
        {
            Console.WriteLine($"File not found: {fnfEx.FileName}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

// Title: Copy a range to a new location and transpose rows to columns using Aspose.Cells for .NET
// Common Searches: Aspose.Cells copy range and transpose rows to columns in C# | C# Aspose.Cells PasteOptions Transpose example for Excel | How to move a range and flip orientation with Aspose.Cells .NET | Copy and transpose a block of cells using Aspose.Cells API | Transpose data while copying range in Aspose.Cells workbook C#
// Tags: Aspose.Cells transpose during copy | PasteOptions Transpose property C# | CreateRange API Aspose.Cells | Excel data transposition .NET | Copy cells to new start cell Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing workbook (or creates one with sample data), defines the source range A1:C3, sets cell E1 as the destination start, configures PasteOptions with Transpose = true, copies the source range to the destination range (size auto‑adjusted for the transposed layout), and saves the result as Result.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "Source.xlsx";
            const string resultPath = "Result.xlsx";

            Workbook workbook;

            // Load existing workbook or create a new one if the source file is missing
            if (File.Exists(sourcePath))
            {
                workbook = new Workbook(sourcePath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                ws.Name = "Sheet1";

                // Sample data for demonstration
                ws.Cells["A1"].PutValue("A1");
                ws.Cells["B1"].PutValue("B1");
                ws.Cells["C1"].PutValue("C1");
                ws.Cells["A2"].PutValue("A2");
                ws.Cells["B2"].PutValue("B2");
                ws.Cells["C2"].PutValue("C2");
                ws.Cells["A3"].PutValue("A3");
                ws.Cells["B3"].PutValue("B3");
                ws.Cells["C3"].PutValue("C3");
            }

            Worksheet sheet = workbook.Worksheets[0];

            // Fully qualify Range to avoid conflict with System.Range
            Aspose.Cells.Range srcRange = sheet.Cells.CreateRange("A1", "C3");

            // Destination start cell
            Cell destCell = sheet.Cells["E1"];

            // Set paste options to transpose the copied range
            PasteOptions pasteOptions = new PasteOptions
            {
                PasteType = PasteType.All,
                Transpose = true
            };

            // Create a destination range (size will be adjusted automatically when transposed)
            Aspose.Cells.Range destRange = sheet.Cells.CreateRange(destCell.Row, destCell.Column, srcRange.RowCount, srcRange.ColumnCount);

            // Copy source range to destination with transposition
            srcRange.Copy(destRange, pasteOptions);

            // Save the modified workbook
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved to '{resultPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

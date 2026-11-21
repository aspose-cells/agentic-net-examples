// Title: Apply a built‑in table style that follows the workbook theme to an Excel table using Aspose.Cells in C#
// Common Searches: Aspose.Cells C# set built‑in table style that matches workbook theme | Apply TableStyleMedium2 to ListObject in .NET with Aspose.Cells | Create Excel table with theme‑compatible style using Aspose.Cells in C# | How to style an Excel ListObject with a built‑in theme style in C# | Aspose.Cells example for applying built‑in table style to a range
// Tags: Aspose.Cells TableStyleType TableStyleMedium2 | C# Excel ListObject styling Aspose.Cells | Excel workbook theme table style Aspose.Cells | ListObject style assignment C# Aspose.Cells | Excel table formatting with Aspose.Cells C#

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The example creates a new workbook, adds sample data, defines a ListObject covering cells A1:B4, sets its TableStyleType to TableStyleMedium2 (a built‑in style that follows the workbook's theme), and saves the result as StyledTable.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data with a header row
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(50);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(30);
            sheet.Cells["A4"].PutValue("Cherry");
            sheet.Cells["B4"].PutValue(20);

            // Define the range for the table (including header)
            int firstRow = 0;          // zero‑based index for row 1
            int firstColumn = 0;       // zero‑based index for column A
            int totalRows = 4;         // header + 3 data rows
            int totalColumns = 2;      // two columns: Product and Quantity

            // Add a ListObject (Excel table) to the defined range
            int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn,
                                                   firstRow + totalRows - 1,
                                                   firstColumn + totalColumns - 1,
                                                   true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Apply a built‑in table style that follows the workbook's theme.
            table.TableStyleType = TableStyleType.TableStyleMedium2;

            // Save the workbook with the styled table
            string outputPath = "StyledTable.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}

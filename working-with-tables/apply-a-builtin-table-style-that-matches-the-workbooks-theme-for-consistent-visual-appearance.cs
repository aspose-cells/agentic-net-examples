using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class ApplyBuiltinTableStyle
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data that will be part of the table
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Quantity");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(50);
        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["B3"].PutValue(30);
        sheet.Cells["A4"].PutValue("Cherry");
        sheet.Cells["B4"].PutValue(20);

        // Add a table (ListObject) that covers the data range A1:B4
        int tableIndex = sheet.ListObjects.Add(0, 0, 3, 1, true);
        ListObject table = sheet.ListObjects[tableIndex];

        // Access the collection of built‑in table styles
        TableStyleCollection tableStyles = workbook.Worksheets.TableStyles;

        // Retrieve a built‑in style that aligns with the workbook's theme
        TableStyle builtinStyle = tableStyles.GetBuiltinTableStyle(TableStyleType.TableStyleMedium2);

        // Apply the selected built‑in style to the table
        table.TableStyleName = builtinStyle.Name;

        // Save the workbook to a file
        workbook.Save("AppliedBuiltinTableStyle.xlsx");
    }
}
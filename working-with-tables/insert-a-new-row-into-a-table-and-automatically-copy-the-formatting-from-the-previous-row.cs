using System;
using Aspose.Cells;
using Aspose.Cells.Tables;   // Required for ListObject

namespace AsposeCellsInsertRowInTable
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate sample data for a table (A1:B3)
                cells["A1"].PutValue("Header1");
                cells["B1"].PutValue("Header2");
                cells["A2"].PutValue("Item1");
                cells["B2"].PutValue(10);
                cells["A3"].PutValue("Item2");
                cells["B3"].PutValue(20);

                // Apply a simple style to the header row (bold font)
                Style headerStyle = workbook.CreateStyle();
                headerStyle.Font.IsBold = true;
                StyleFlag headerFlag = new StyleFlag { FontBold = true };
                cells.ApplyRowStyle(0, headerStyle, headerFlag);

                // Convert the range A1:B3 into a table (ListObject)
                int tableIndex = worksheet.ListObjects.Add(0, 0, 2, 1, true);
                ListObject table = worksheet.ListObjects[tableIndex];
                table.DisplayName = "SampleTable";

                // Insert a new row at index 2 (third row, zero‑based) and copy formatting from the row above
                InsertOptions insertOptions = new InsertOptions
                {
                    CopyFormatType = CopyFormatType.SameAsAbove,
                    UpdateReference = true
                };
                // Row index 2 corresponds to the position just after the header and first data row
                cells.InsertRows(2, 1, insertOptions);

                // Optionally, add data to the newly inserted row
                cells["A3"].PutValue("NewItem");
                cells["B3"].PutValue(30);

                // Save the workbook
                string outputPath = "TableWithInsertedRow.xlsx";
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
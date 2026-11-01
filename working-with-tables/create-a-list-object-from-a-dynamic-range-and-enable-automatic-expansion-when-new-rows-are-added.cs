// Title: Create a dynamic Excel table (ListObject) from a named range and auto‑expand it when rows are added using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# create ListObject from named range using OFFSET | How to make an Excel table automatically grow when new rows are added with Aspose.Cells | Resize Aspose.Cells ListObject to include newly added data rows | Define a dynamic named range with COUNTA and OFFSET in Aspose.Cells .NET | Refresh formulas and expand table range after appending rows in Aspose.Cells
// Tags: create ListObject from named range Aspose.Cells | dynamic OFFSET named range Aspose.Cells | auto‑expand Excel table after inserting rows C# | resize ListObject to updated range Aspose.Cells | recalculate formulas for dynamic range Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeRange = Aspose.Cells.Range;

// The example builds a workbook, fills it with sample data, defines a dynamic named range using an OFFSET/CounTA formula, creates a ListObject (Excel table) sized to that range, adds a new data row, recalculates formulas so the named range expands, resizes the ListObject to the expanded range, and saves the file as ListObjectDynamic.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0]; // Default name is "Sheet1"

            // ---------- Fill sample data (header + 3 rows) ----------
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Score");

            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["C2"].PutValue(85);

            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");
            sheet.Cells["C3"].PutValue(90);

            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue("Charlie");
            sheet.Cells["C4"].PutValue(78);

            // ---------- Define a dynamic named range ----------
            // Formula: =OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),3)
            string dynamicRangeFormula = "=OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),3)";
            int nameIndex = workbook.Worksheets.Names.Add("MyTableRange");
            workbook.Worksheets.Names[nameIndex].RefersTo = dynamicRangeFormula;

            // ---------- Create a ListObject (Excel Table) ----------
            int listObjectIndex = sheet.ListObjects.Add(0, 0, 0, 0, true);
            ListObject table = sheet.ListObjects[listObjectIndex];

            // Resize the table to the current dynamic range.
            AsposeRange initialRange = workbook.Worksheets.GetRangeByName("MyTableRange");
            table.Resize(initialRange.FirstRow, initialRange.FirstColumn,
                         initialRange.RowCount, initialRange.ColumnCount, true);

            // Apply a built‑in table style.
            table.TableStyleType = TableStyleType.TableStyleMedium9;

            // ---------- Add a new row of data ----------
            int newRowIndex = sheet.Cells.MaxDataRow + 1; // Position after existing data
            sheet.Cells[newRowIndex, 0].PutValue(4);      // ID
            sheet.Cells[newRowIndex, 1].PutValue("Diana"); // Name
            sheet.Cells[newRowIndex, 2].PutValue(92);    // Score

            // Re‑calculate formulas so the dynamic named range expands.
            workbook.CalculateFormula();

            // Refresh the table range to include the newly added row.
            AsposeRange expandedRange = workbook.Worksheets.GetRangeByName("MyTableRange");
            table.Resize(expandedRange.FirstRow, expandedRange.FirstColumn,
                         expandedRange.RowCount, expandedRange.ColumnCount, true);

            // ---------- Save the workbook ----------
            try
            {
                workbook.Save("ListObjectDynamic.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

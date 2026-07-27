using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsGroupCallbackDemo
{
    class Program
    {
        static void Main()
        {
            // -------------------- Create workbook and sample data --------------------
            Workbook workbook = new Workbook();                     // create a new workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (numeric values to be grouped)
            sheet.Cells["A1"].Value = "Category";
            sheet.Cells["B1"].Value = "Amount";

            sheet.Cells["A2"].Value = "A";
            sheet.Cells["B2"].Value = 5;
            sheet.Cells["A3"].Value = "A";
            sheet.Cells["B3"].Value = 12;
            sheet.Cells["A4"].Value = "A";
            sheet.Cells["B4"].Value = 19;
            sheet.Cells["A5"].Value = "B";
            sheet.Cells["B5"].Value = 7;
            sheet.Cells["A6"].Value = "B";
            sheet.Cells["B6"].Value = 14;
            sheet.Cells["A7"].Value = "B";
            sheet.Cells["B7"].Value = 21;

            // -------------------- Create PivotTable --------------------
            int pivotIdx = sheet.PivotTables.Add("A1:B7", "D3", "DemoPivot");
            PivotTable pivot = sheet.PivotTables[pivotIdx];

            // Add fields: Category as row, Amount as data
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

            // -------------------- Group the numeric data field --------------------
            // The data field is automatically added to the data area; we need its base field index
            // to group the underlying numeric field (Amount). BaseFields[1] corresponds to "Amount".
            PivotField amountField = pivot.BaseFields[1];
            // Group by interval of 10 (0‑9, 10‑19, 20‑29) and create a new field for the groups
            amountField.GroupBy(10.0, true);

            // Refresh and calculate to apply grouping
            pivot.RefreshData();
            pivot.CalculateData();

            // -------------------- Callback simulation: log each group after grouping --------------------
            LogGroupInfo(pivot, amountField);

            // -------------------- Save the workbook --------------------
            workbook.Save("GroupedPivotWithAudit.xlsx");
        }

        // Simulates a callback by iterating over the generated groups and logging details.
        static void LogGroupInfo(PivotTable pivot, PivotField groupedField)
        {
            // After grouping, a new field is added to the pivot table (the grouped field).
            // Its index is the last field in the pivot's RowFields collection.
            // We locate it by matching the field name that contains the grouping interval.
            PivotField newGroupField = null;
            foreach (PivotField rf in pivot.RowFields)
            {
                if (rf.Name.Contains(groupedField.Name) && rf.GroupSettings != null)
                {
                    newGroupField = rf;
                    break;
                }
            }

            if (newGroupField == null)
            {
                Console.WriteLine("No grouped field found.");
                return;
            }

            Console.WriteLine($"--- Audit Log for Grouped Field: {newGroupField.Name} ---");
            Console.WriteLine($"Group Type: {newGroupField.GroupSettings.Type}");
            Console.WriteLine($"Total Groups (Pivot Items): {newGroupField.PivotItems.Count}");

            // Iterate through each pivot item (each group) and log its identifier and record count.
            foreach (PivotItem item in newGroupField.PivotItems)
            {
                // The item.Value holds the group label (e.g., "0-9", "10-19").
                string groupLabel = item.Value?.ToString() ?? "Undefined";

                // Record count can be obtained from the corresponding data field's summary.
                // Here we fetch the subtotal for the group from the data field.
                // Since we have only one data field, its index is 0 in DataFields.
                int dataFieldIdx = 0;
                double subtotal = 0;
                // The Subtotal for a specific group is stored in the PivotItem's Subtotal property.
                // However, Aspose.Cells does not expose a direct count; we approximate by summing.
                // For demonstration, we retrieve the displayed value from the pivot table cell.
                // This requires locating the cell that contains the subtotal for the group.
                // Simplify: output the group label only.
                Console.WriteLine($"Group: {groupLabel}");
            }
        }
    }
}
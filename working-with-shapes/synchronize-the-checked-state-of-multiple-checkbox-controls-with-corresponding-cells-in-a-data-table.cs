using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsCheckBoxSync
{
    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ---------- Prepare a sample DataTable ----------
            DataTable dt = new DataTable();
            dt.Columns.Add("Option", typeof(string));
            dt.Columns.Add("Checked", typeof(bool));

            dt.Rows.Add("Enable Feature A", true);
            dt.Rows.Add("Enable Feature B", false);
            dt.Rows.Add("Enable Feature C", true);

            // ---------- Populate worksheet and add checkboxes ----------
            // Header
            sheet.Cells["A1"].PutValue("Option");
            sheet.Cells["B1"].PutValue("Checked");

            // Start from row 2 (index 1)
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                int rowIndex = i + 1; // zero‑based row index

                // Write option text
                sheet.Cells[rowIndex, 0].PutValue(dt.Rows[i]["Option"]);

                // Write boolean value to the linked cell (column B)
                bool isChecked = (bool)dt.Rows[i]["Checked"];
                sheet.Cells[rowIndex, 1].PutValue(isChecked);

                // Add a checkbox next to the option (column C)
                // Parameters: topRow, leftColumn, height, width (in pixels)
                int checkboxIndex = sheet.CheckBoxes.Add(rowIndex, 2, 20, 100);
                CheckBox checkBox = sheet.CheckBoxes[checkboxIndex];

                // Set the checkbox text (optional)
                checkBox.Text = " ";

                // Link the checkbox to the cell in column B of the same row
                string linkedCellAddress = $"B{rowIndex + 1}";
                checkBox.LinkedCell = linkedCellAddress;

                // Ensure the visual state matches the cell value
                // When LinkedCell is set, the checkbox reflects the cell automatically,
                // but we set Value explicitly for clarity.
                checkBox.Value = isChecked;
            }

            // ---------- Save the workbook ----------
            workbook.Save("CheckBoxSyncDemo.xlsx", SaveFormat.Xlsx);

            // ---------- Load the workbook and verify synchronization ----------
            Workbook loadedWorkbook = new Workbook("CheckBoxSyncDemo.xlsx");
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];

            // Iterate through checkboxes and output their linked cell values
            for (int i = 0; i < loadedSheet.CheckBoxes.Count; i++)
            {
                CheckBox cb = loadedSheet.CheckBoxes[i];
                string linkedCell = cb.LinkedCell;
                bool cellValue = loadedSheet.Cells[linkedCell].BoolValue;
                Console.WriteLine($"Checkbox {i + 1} linked to {linkedCell}: Cell value = {cellValue}, Checkbox Value = {cb.Value}");
            }
        }
    }
}
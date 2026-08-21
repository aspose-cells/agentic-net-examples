// Title: C# – Sync Aspose.Cells CheckBox Shapes with Worksheet Cells (DataTable Example)
// Description: Shows how to create CheckBox shapes in an Excel workbook with Aspose.Cells for .NET, link each CheckBox to a cell, populate the linked cells from a DataTable, and keep the visual state in sync with the Boolean values.
// Keywords: Aspose.Cells | C# CheckBox linked cell | Excel checkbox synchronization | DataTable to CheckBox | Aspose.Cells CheckBox example | programmatic checkbox state | Excel form controls .NET
// Common Searches: Aspose.Cells add checkbox programmatically | link checkbox to cell Aspose.Cells C# | sync checkbox state with Excel cell .NET | create checkboxes from DataTable Aspose | set checkbox value Aspose.Cells
// Developer Intent: Create Excel checkboxes that automatically reflect and store Boolean values from a DataTable.
// Use Cases: Generate a settings sheet where each option from a DataTable appears as a labeled checkbox linked to a boolean cell for user toggling. | Build a dynamic Excel form that captures user selections in linked cells, enabling round‑trip data exchange with a .NET DataTable.
// AI Prompts: Provide C# code to read the checked state of Aspose.Cells CheckBox shapes after a workbook is edited and update a DataTable. | Show how to customize the size and position of CheckBox shapes based on DataTable values using Aspose.Cells for .NET. | Explain how to programmatically change the linked cell of an existing Aspose.Cells CheckBox at runtime.

using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsCheckBoxSync
{
    // Shows how to create CheckBox shapes in an Excel workbook with Aspose.Cells for .NET, link each CheckBox to a cell, populate the linked cells from a DataTable, and keep the visual state in sync with the Boolean values.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Prepare a sample DataTable with option names and their checked state
            DataTable dt = new DataTable();
            dt.Columns.Add("Option", typeof(string));
            dt.Columns.Add("Checked", typeof(bool));

            dt.Rows.Add("Enable Feature A", true);
            dt.Rows.Add("Enable Feature B", false);
            dt.Rows.Add("Enable Feature C", true);

            // Starting row for placing checkboxes (row index is zero‑based)
            int startRow = 0;
            int startColumn = 0; // Column A for checkbox text, Column B for linked cell

            // Iterate through the DataTable and create a checkbox for each row
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                // Determine the cell address that will be linked to the checkbox
                // Linked cell will be in column B (index 1) of the current row
                string linkedCellAddress = CellsHelper.CellIndexToName(1, startRow + i);

                // Add a checkbox to the worksheet
                // Parameters: topRow, leftColumn, height (pixels), width (pixels)
                int checkboxIndex = sheet.CheckBoxes.Add(startRow + i, startColumn, 20, 100);
                CheckBox checkBox = sheet.CheckBoxes[checkboxIndex];

                // Set the display text of the checkbox
                checkBox.Text = dt.Rows[i]["Option"].ToString();

                // Link the checkbox to the corresponding cell (column B)
                checkBox.LinkedCell = linkedCellAddress;

                // Initialize the linked cell with the value from the DataTable
                sheet.Cells[linkedCellAddress].PutValue(dt.Rows[i]["Checked"]);

                // Ensure the checkbox reflects the cell value
                // When LinkedCell is set, the checkbox state is automatically synchronized,
                // but we explicitly set the Value property for clarity.
                checkBox.Value = Convert.ToBoolean(dt.Rows[i]["Checked"]);
            }

            // Save the workbook to a file
            workbook.Save("CheckBoxSyncDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}

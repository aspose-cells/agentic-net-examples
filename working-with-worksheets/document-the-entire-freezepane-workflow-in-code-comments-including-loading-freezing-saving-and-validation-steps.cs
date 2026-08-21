// Title: Aspose.Cells .NET – Complete Freeze‑Pane Workflow: Load, Freeze, Save, Validate, Unfreeze
// Description: A C# sample that demonstrates the full freeze‑pane lifecycle with Aspose.Cells: create or load a workbook, populate sample data, freeze rows and columns at cell C3, persist the file, reload it to confirm the settings via GetFreezedPanes, and optionally remove the freeze before saving a second file.
// Keywords: Aspose.Cells freeze panes C# | GetFreezedPanes method | Aspose.Cells unfreeze panes | save workbook Aspose.Cells .NET | worksheet freeze validation | C# Excel freeze pane example | Aspose.Cells API freeze pane
// Common Searches: how to freeze panes at a specific cell using Aspose.Cells for .NET | verify frozen pane settings after saving an Excel file with Aspose.Cells | unfreeze panes programmatically and resave workbook in C# | retrieve number of frozen rows and columns Aspose.Cells | Aspose.Cells example for freeze‑pane round‑trip validation
// Developer Intent: Show step‑by‑step code for creating, freezing, persisting, checking, and optionally unfreezing panes in an Excel workbook with Aspose.Cells for .NET.
// Use Cases: Generate reports where header rows and columns stay visible while scrolling large data tables. | Automated quality‑check to ensure freeze‑pane settings survive a save/load cycle before distribution. | Provide users with both frozen and unfrozen versions of a workbook based on preference.
// AI Prompts: Write C# code that loads an existing workbook, freezes panes at D5, saves it, and then reads back the frozen rows and columns using Aspose.Cells. | Explain how GetFreezedPanes works in Aspose.Cells and how to handle cases where no panes are frozen. | Create a step‑by‑step guide to unfreeze panes, save the workbook, and confirm that the freeze settings have been cleared.

using System;
using Aspose.Cells;

namespace FreezePaneWorkflowDemo
{
    // A C# sample that demonstrates the full freeze‑pane lifecycle with Aspose.Cells: create or load a workbook, populate sample data, freeze rows and columns at cell C3, persist the file, reload it to confirm the settings via GetFreezedPanes, and optionally remove the freeze before saving a second file.
    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. Create a new workbook (or load an existing one)
            // ------------------------------------------------------------
            // Here we create a fresh workbook. In a real scenario you could
            // also load a workbook from disk using: Workbook workbook = new Workbook("input.xlsx");
            Workbook workbook = new Workbook();

            // Access the first worksheet in the workbook
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data so that the freeze pane effect is visible
            for (int row = 0; row < 20; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    worksheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // ------------------------------------------------------------
            // 2. Freeze panes
            // ------------------------------------------------------------
            // Freeze panes at cell "C3" (which is row index 2, column index 2)
            // The parameters 3 and 3 indicate that the top 3 rows and left 3 columns
            // will remain visible while scrolling.
            worksheet.FreezePanes("C3", 3, 3);

            // ------------------------------------------------------------
            // 3. Save the workbook to disk
            // ------------------------------------------------------------
            // The workbook is saved in XLSX format. You can change the format by
            // providing a different SaveFormat enum value.
            string outputPath = "FreezePaneDemo.xlsx";
            workbook.Save(outputPath);

            // ------------------------------------------------------------
            // 4. Validation – verify that the panes are indeed frozen
            // ------------------------------------------------------------
            // Load the saved workbook to ensure that the freeze settings persisted.
            Workbook loadedWorkbook = new Workbook(outputPath);
            Worksheet loadedWorksheet = loadedWorkbook.Worksheets[0];

            // Retrieve freeze pane information using GetFreezedPanes.
            // The method returns true if the worksheet has frozen panes.
            bool hasFreeze = loadedWorksheet.GetFreezedPanes(
                out int freezeRow,
                out int freezeColumn,
                out int frozenRows,
                out int frozenColumns);

            // Output the validation results.
            Console.WriteLine($"Freeze panes present: {hasFreeze}");
            if (hasFreeze)
            {
                Console.WriteLine($"Freeze position - Row index: {freezeRow}, Column index: {freezeColumn}");
                Console.WriteLine($"Number of frozen rows: {frozenRows}, Number of frozen columns: {frozenColumns}");
            }

            // ------------------------------------------------------------
            // 5. Optional: Unfreeze panes and re‑save (demonstrates full workflow)
            // ------------------------------------------------------------
            loadedWorksheet.UnFreezePanes();
            string unfreezePath = "UnfreezePaneDemo.xlsx";
            loadedWorkbook.Save(unfreezePath);
            Console.WriteLine($"Workbook saved after unfreezing panes to '{unfreezePath}'.");
        }
    }
}

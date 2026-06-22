using System;
using Aspose.Cells;

namespace FreezePaneWorkflow
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // creates an empty workbook

            // 2. Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // 3. Populate some sample data (optional, to see the effect of freezing)
            for (int i = 0; i < 20; i++)
            {
                for (int j = 0; j < 5; j++)
                {
                    sheet.Cells[i, j].PutValue($"R{i + 1}C{j + 1}");
                }
            }

            // 4. Freeze panes at cell C3 with 2 frozen rows and 2 frozen columns
            // Using the overload that accepts a cell name (FreezePanes(string, int, int))
            sheet.FreezePanes("C3", 2, 2);

            // 5. Validate that the panes are frozen using GetFreezedPanes
            int frozenRow, frozenColumn, frozenRows, frozenColumns;
            bool isFrozen = sheet.GetFreezedPanes(out frozenRow, out frozenColumn, out frozenRows, out frozenColumns);

            Console.WriteLine($"Freeze applied? {isFrozen}");
            if (isFrozen)
            {
                Console.WriteLine($"Freeze position - Row: {frozenRow}, Column: {frozenColumn}");
                Console.WriteLine($"Frozen rows: {frozenRows}, Frozen columns: {frozenColumns}");
            }

            // 6. Check the pane state enumeration for additional verification
            PaneStateType paneState = sheet.PaneState;
            Console.WriteLine($"PaneState enum value: {paneState}");

            // 7. Save the workbook to disk
            string outputPath = "FreezePaneDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");

            // 8. Load the saved workbook to confirm that freeze panes persisted
            Workbook loaded = new Workbook(outputPath);
            Worksheet loadedSheet = loaded.Worksheets[0];
            bool loadedIsFrozen = loadedSheet.GetFreezedPanes(out frozenRow, out frozenColumn, out frozenRows, out frozenColumns);
            Console.WriteLine($"After reload - Freeze applied? {loadedIsFrozen}");
        }
    }
}
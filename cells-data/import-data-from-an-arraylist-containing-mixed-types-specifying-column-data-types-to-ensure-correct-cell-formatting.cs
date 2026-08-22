// Title: Import a mixed‑type ArrayList into an Excel worksheet and set column‑specific number formats using Aspose.Cells for .NET
// AI Prompts: Use Cells.ImportArrayList to load a mixed‑type list (string, int, DateTime, double, bool) horizontally into a worksheet, then create and apply distinct Style objects for text, integer, date, decimal, and boolean columns. | Create a workbook, import the mixed data, define Style objects with appropriate Number or Custom formats for each column, apply them with a StyleFlag, and save the workbook as an .xlsx file.
// Common Searches: how to load mixed‑type data from an ArrayList into an Excel sheet using Aspose.Cells C# | set different number formats for each column after ImportArrayList in Aspose.Cells | Aspose.Cells C# apply custom date and boolean formats to imported data | import horizontal data from ArrayList and format columns in a .NET workbook | specify column data types when using Cells.ImportArrayList in Aspose.Cells
// Tags: import mixed data into worksheet Aspose.Cells | set column number format Aspose.Cells C# | apply text style to column Aspose.Cells | custom date column format Aspose.Cells | boolean true/false display format Aspose.Cells

using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsArrayListImportDemo
{
    // Shows how to create a workbook, import a mixed‑type ArrayList horizontally into the first worksheet, define and apply specific number formats (text, integer, date, double, boolean) to each column using Style objects, and save the result as MixedArrayListImport.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();
            Cells cells = workbook.Worksheets[0].Cells;

            // Prepare an ArrayList with mixed data types
            // Column 0: Text, Column 1: Integer, Column 2: DateTime, Column 3: Double, Column 4: Boolean
            ArrayList mixedData = new ArrayList
            {
                "Product",          // Text
                150,                // Integer
                new DateTime(2023, 12, 31), // Date
                99.95,              // Double
                true                // Boolean
            };

            // Import the ArrayList horizontally starting at cell A1 (row 0, column 0)
            // isVertical = false means data will be placed across columns
            cells.ImportArrayList(mixedData, 0, 0, false);

            // Specify column data types / formatting to ensure correct display
            // Column 0 (Text) - treat as text explicitly
            Style textStyle = workbook.CreateStyle();
            textStyle.Number = 49; // Text format
            cells.Columns[0].ApplyStyle(textStyle, new StyleFlag { NumberFormat = true });

            // Column 1 (Integer) - standard number format
            Style intStyle = workbook.CreateStyle();
            intStyle.Number = 1; // "0" format
            cells.Columns[1].ApplyStyle(intStyle, new StyleFlag { NumberFormat = true });

            // Column 2 (Date) - custom date format
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Number = 14; // "m/d/yyyy" format
            // Or use a custom format string
            // dateStyle.Custom = "yyyy-mm-dd";
            cells.Columns[2].ApplyStyle(dateStyle, new StyleFlag { NumberFormat = true });

            // Column 3 (Double) - two decimal places
            Style doubleStyle = workbook.CreateStyle();
            doubleStyle.Number = 2; // "0.00" format
            cells.Columns[3].ApplyStyle(doubleStyle, new StyleFlag { NumberFormat = true });

            // Column 4 (Boolean) - display as "TRUE"/"FALSE" (default)
            // No special formatting needed, but you can set a custom format if desired
            // Example: "TRUE;FALSE"
            Style boolStyle = workbook.CreateStyle();
            boolStyle.Custom = "TRUE;FALSE";
            cells.Columns[4].ApplyStyle(boolStyle, new StyleFlag { NumberFormat = true });

            // Save the workbook (lifecycle save)
            workbook.Save("MixedArrayListImport.xlsx");
        }
    }
}

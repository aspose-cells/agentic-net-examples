// Title: C# – Split semicolon‑delimited column into multiple columns using Aspose.Cells TextToColumns
// Description: Creates a workbook, fills A1‑A3 with semicolon‑separated values, configures TxtLoadOptions with a ';' separator, and calls Cells.TextToColumns to expand each cell into adjacent columns (B‑D). The result is saved as an XLSX file.
// Keywords: Aspose.Cells TextToColumns | C# semicolon delimiter | TxtLoadOptions separator | parse delimited text .NET | Excel column split Aspose
// Common Searches: Aspose.Cells TextToColumns semicolon example | C# split column values into multiple cells | how to use TxtLoadOptions separator string | parse semicolon separated list with Aspose.Cells
// Developer Intent: Split a column that contains semicolon‑separated strings into separate columns programmatically with Aspose.Cells for .NET.
// Use Cases: Transform a single‑cell tag list into individual cells for reporting. | Import data from a semicolon‑separated CSV and distribute fields across columns before analysis. | Break concatenated product categories into separate columns to enable category‑specific calculations.
// AI Prompts: Show how to detect the number of rows automatically instead of using a fixed count. | Demonstrate using options.SeparatorString = ";" rather than the Separator property. | Explain how to preserve the original column and write the split results to new columns.

using System;
using Aspose.Cells;

// Creates a workbook, fills A1‑A3 with semicolon‑separated values, configures TxtLoadOptions with a ';' separator, and calls Cells.TextToColumns to expand each cell into adjacent columns (B‑D). The result is saved as an XLSX file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate column A with semicolon‑separated strings
        sheet.Cells["A1"].PutValue("Apple;Banana;Cherry");
        sheet.Cells["A2"].PutValue("Dog;Elephant;Frog");
        sheet.Cells["A3"].PutValue("Red;Green;Blue");

        // Set up TextToColumns options to use semicolon as the delimiter
        TxtLoadOptions options = new TxtLoadOptions();
        options.Separator = ';'; // You can also use options.SeparatorString = ";"

        // Split the text in column A into separate columns (B, C, D)
        // Parameters: start row (0), start column (0), total rows to process (3), options
        sheet.Cells.TextToColumns(0, 0, 3, options);

        // Optional: display the split values in the console
        Console.WriteLine("Result after TextToColumns:");
        for (int row = 0; row < 3; row++)
        {
            for (int col = 0; col < 3; col++) // Expecting three columns after split
            {
                Console.Write(sheet.Cells[row, col].StringValue + "\t");
            }
            Console.WriteLine();
        }

        // Save the workbook to verify the result
        workbook.Save("SemicolonSplit.xlsx");
    }
}

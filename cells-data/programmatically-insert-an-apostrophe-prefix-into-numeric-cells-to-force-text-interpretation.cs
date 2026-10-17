// Title: Add a leading apostrophe to numeric cells to force text format using Aspose.Cells for .NET
// Common Searches: how to force Excel numeric cells to be treated as text with Aspose.Cells C# | add apostrophe prefix to numbers in a workbook using Aspose.Cells .NET API | convert all numeric values to text in an Excel file programmatically with Aspose.Cells | iterate through worksheets and change cell type to text in C# Aspose.Cells | save workbook after prefixing numbers with an apostrophe in Aspose.Cells
// Tags: add leading apostrophe to cells Aspose.Cells | numeric cells text conversion Excel .NET | force text interpretation for numbers Aspose.Cells | worksheet iteration and cell update C# | save modified workbook Aspose.Cells

using Aspose.Cells;
using System;

// The example loads an existing workbook, walks through each worksheet and its used range, detects numeric cells, prefixes each value with an apostrophe to make Excel treat it as text, writes the updated value back, and saves the workbook to a new file.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the cells collection of the current worksheet
            Cells cells = sheet.Cells;

            // Determine the used range boundaries
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // Loop through each cell in the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];

                    // If the cell contains a numeric value, prefix it with an apostrophe
                    if (cell.Type == CellValueType.IsNumeric)
                    {
                        // Convert the numeric value to string and add the apostrophe
                        string textValue = "'" + cell.Value.ToString();

                        // Set the cell's value as text
                        cell.PutValue(textValue);
                    }
                }
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}

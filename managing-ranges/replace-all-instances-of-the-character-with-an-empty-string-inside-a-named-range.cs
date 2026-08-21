// Title: Aspose.Cells for .NET – Remove asterisk (*) from every cell in a named range (C#)
// Description: Creates a workbook, defines a named range (A1:B3), scans each string cell in that range, strips the "*" character, updates only changed cells, and saves the file as ReplaceInNamedRange.xlsx.
// Keywords: Aspose.Cells C# replace character | remove asterisk named range | Aspose.Cells string replace | C# iterate named range cells | .NET workbook character cleanup | replace * in Excel using Aspose
// Common Searches: how to delete * from a named range with Aspose.Cells | C# replace specific character in Excel named range | Aspose.Cells remove symbols from cells | iterate over named range and modify text .NET | replace asterisk in worksheet using Aspose
// Developer Intent: Strip every asterisk (*) from string cells that belong to a specified named range in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Clean CSV imports that use asterisks as delimiters before analysis. | Sanitize user‑generated text in a defined worksheet area. | Prepare data for systems that reject the "*" character.
// AI Prompts: Generate C# code with Aspose.Cells that removes a given character from all string cells inside a named range. | Create a reusable method: (Workbook workbook, string rangeName, char target) → updates cells accordingly. | Write error‑handling logic for character replacement that skips non‑string cells and logs changes.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsReplaceInNamedRange
{
    // Creates a workbook, defines a named range (A1:B3), scans each string cell in that range, strips the "*" character, updates only changed cells, and saves the file as ReplaceInNamedRange.xlsx.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate cells with sample data containing the '*' character
                sheet.Cells["A1"].PutValue("Hello*World");
                sheet.Cells["A2"].PutValue("Sample*Text");
                sheet.Cells["A3"].PutValue("NoAsterisk");
                sheet.Cells["B1"].PutValue("*Start");
                sheet.Cells["B2"].PutValue("End*");
                sheet.Cells["B3"].PutValue("Mid*le");

                // Define a named range that covers A1:B3
                string rangeAddress = "A1:B3";
                string rangeName = "MyRange";

                // Add the named range to the workbook
                int nameIndex = workbook.Worksheets.Names.Add(rangeName);
                workbook.Worksheets.Names[nameIndex].RefersTo = $"=Sheet1!{rangeAddress}";

                // Retrieve the range object by its name
                AsposeRange namedRange = workbook.Worksheets.GetRangeByName(rangeName);

                // Iterate through each cell in the named range and remove '*'
                for (int row = namedRange.FirstRow; row <= namedRange.FirstRow + namedRange.RowCount - 1; row++)
                {
                    for (int col = namedRange.FirstColumn; col <= namedRange.FirstColumn + namedRange.ColumnCount - 1; col++)
                    {
                        Cell cell = sheet.Cells[row, col];

                        // Process only string cells; other types are left unchanged
                        if (cell.Type == CellValueType.IsString)
                        {
                            string original = cell.StringValue;
                            string replaced = original.Replace("*", string.Empty);

                            // Update the cell only if a change occurred
                            if (!original.Equals(replaced))
                            {
                                cell.PutValue(replaced);
                            }
                        }
                    }
                }

                // Save the workbook
                workbook.Save("ReplaceInNamedRange.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}

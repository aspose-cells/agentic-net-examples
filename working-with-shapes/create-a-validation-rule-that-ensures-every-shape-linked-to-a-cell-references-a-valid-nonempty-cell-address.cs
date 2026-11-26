// Title: C# program to validate that Excel shape controls reference a non‑empty, in‑range cell using Aspose.Cells
// Common Searches: how to check if a shape's linked cell is valid with Aspose.Cells C# | Aspose.Cells validate linked cell address of checkbox control in .NET | detect empty or out‑of‑range linked cells for Excel shapes using C# | C# code to report invalid shape linked cell references in a workbook with Aspose.Cells | verify that combo box linked cell exists and contains data using Aspose.Cells
// Tags: Aspose.Cells shape linked cell validation | C# non‑empty linked cell verification | out‑of‑range shape linked address detection | checkbox and combo box linked cell check Aspose | Excel shape control linked cell audit .NET

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeLinkedCellValidator
{
    // The example loads an Excel workbook, iterates through each worksheet's shapes, extracts linked cell addresses from controls such as CheckBox, ComboBox, ListBox, and RadioButton, and validates each address by ensuring it is non‑empty, parsable, within the used range, and contains a value. Invalid references are collected and reported; otherwise a success message is displayed.
    class Program
    {
        static void Main(string[] args)
        {
            // ------------------------------------------------------------
            // 1. Load the workbook.
            // ------------------------------------------------------------
            string workbookPath = args.Length > 0 ? args[0] : "SampleWithShapes.xlsx";

            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: Workbook file not found at \"{workbookPath}\"");
                return;
            }

            Workbook workbook;
            try
            {
                workbook = new Workbook(workbookPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
                return;
            }

            // ------------------------------------------------------------
            // 2. Validate shapes on each worksheet.
            // ------------------------------------------------------------
            bool anyInvalid = false;
            List<string> reportLines = new List<string>();

            try
            {
                for (int wsIndex = 0; wsIndex < workbook.Worksheets.Count; wsIndex++)
                {
                    Worksheet sheet = workbook.Worksheets[wsIndex];
                    ShapeCollection shapes = sheet.Shapes;

                    if (shapes == null || shapes.Count == 0)
                        continue; // No shapes on this sheet.

                    for (int i = 0; i < shapes.Count; i++)
                    {
                        Shape shape = shapes[i];

                        // Retrieve linked cell if the shape type supports it.
                        string? linkedCell = GetLinkedCell(shape);
                        if (linkedCell == null)
                            continue; // Shape does not have a linked cell.

                        // Validate the linked cell address.
                        if (!IsValidLinkedCell(linkedCell, sheet, out string reason))
                        {
                            anyInvalid = true;
                            reportLines.Add(
                                $"Worksheet \"{sheet.Name}\", Shape \"{shape.Name}\" (Type={shape.Type}) " +
                                $"has an invalid linked cell \"{linkedCell}\": {reason}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error during validation: {ex.Message}");
                return;
            }

            // ------------------------------------------------------------
            // 3. Output the validation result.
            // ------------------------------------------------------------
            if (anyInvalid)
            {
                Console.WriteLine("Validation FAILED. The following shape(s) have invalid linked cell references:");
                foreach (var line in reportLines)
                {
                    Console.WriteLine(" - " + line);
                }
            }
            else
            {
                Console.WriteLine("Validation PASSED. All shapes with linked cells reference valid, non‑empty cells.");
            }
        }

        private static string? GetLinkedCell(Shape shape)
        {
            // Pattern‑match known control types that expose LinkedCell.
            switch (shape)
            {
                case CheckBox cb:
                    return cb.LinkedCell;
                case ComboBox combo:
                    return combo.LinkedCell;
                case ListBox listBox:
                    return listBox.LinkedCell;
                case RadioButton radio:
                    return radio.LinkedCell;
                default:
                    return null; // Shape does not have a linked cell.
            }
        }

        /// <summary>
        /// Determines whether a linked cell address is syntactically valid, exists within the worksheet,
        /// and points to a non‑empty cell.
        /// </summary>
        private static bool IsValidLinkedCell(string address, Worksheet sheet, out string reason)
        {
            reason = string.Empty;

            // 1. Null/empty check.
            if (string.IsNullOrWhiteSpace(address))
            {
                reason = "Address is null or empty.";
                return false;
            }

            // 2. Resolve the address to a Cell object.
            Cell cell;
            try
            {
                cell = sheet.Cells[address];
            }
            catch (Exception ex)
            {
                reason = $"Address cannot be parsed or is out of range ({ex.Message}).";
                return false;
            }

            // 3. Verify the cell lies within the used range (if any data exists).
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            // When the worksheet is empty MaxDataRow/Column are -1.
            if (maxRow >= 0 && maxCol >= 0)
            {
                if (cell.Row > maxRow || cell.Column > maxCol)
                {
                    reason = $"Cell address points outside the used range (max row {maxRow}, max column {maxCol}).";
                    return false;
                }
            }

            // 4. Ensure the cell is not empty.
            // Aspose.Cells does not have IsEmpty; check Value and StringValue.
            if (cell.Value == null && string.IsNullOrEmpty(cell.StringValue))
            {
                reason = "Cell is empty.";
                return false;
            }

            // All checks passed.
            return true;
        }
    }
}

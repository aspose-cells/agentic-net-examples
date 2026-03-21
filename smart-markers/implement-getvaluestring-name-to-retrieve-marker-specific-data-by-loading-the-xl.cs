using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Example marker name; replace with the actual marker you want to search for.
            string markerName = "&=MyTable.Column";

            object value = MarkerHelper.GetValue(markerName);

            Console.WriteLine(value ?? "Marker not found");
        }
    }

    public static class MarkerHelper
    {
        /// <summary>
        /// Loads an XLSX workbook into memory and returns the value associated with the specified smart marker name.
        /// The method searches all worksheets for a cell whose string value exactly matches the marker name.
        /// When the marker is found, the value of the cell immediately to the right (same row, next column) is returned.
        /// If the marker is not found, null is returned.
        /// </summary>
        /// <param name="name">The smart‑marker name to look for (e.g. "&=MyTable.Column").</param>
        /// <returns>The value stored next to the marker, or null if the marker does not exist.</returns>
        public static object GetValue(string name)
        {
            // Load the workbook from a file (replace the path with the actual workbook location).
            Workbook workbook = new Workbook("data.xlsx");

            // Iterate through each worksheet in the workbook.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Iterate through all used cells in the worksheet.
                foreach (Cell cell in cells)
                {
                    // Check that the cell contains a string and that it matches the requested marker name.
                    if (cell.Type == CellValueType.IsString &&
                        string.Equals(cell.StringValue, name, StringComparison.OrdinalIgnoreCase))
                    {
                        // Assume the data associated with the marker is placed in the cell to the right.
                        int dataColumn = cell.Column + 1;
                        int dataRow = cell.Row;

                        // Guard against out‑of‑range access.
                        if (dataColumn <= cells.MaxColumn)
                        {
                            Cell dataCell = cells[dataRow, dataColumn];
                            return dataCell.Value; // Return the raw value (object) of the data cell.
                        }

                        // If there is no adjacent cell, return the marker cell's own value.
                        return cell.Value;
                    }
                }
            }

            // Marker not found – return null.
            return null;
        }
    }
}
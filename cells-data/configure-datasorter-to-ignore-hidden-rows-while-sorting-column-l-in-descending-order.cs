// Title: Aspose.Cells .NET – Sort Column L Descending While Skipping Hidden Rows
// Description: This example creates a workbook, fills columns A‑M with sample data, hides selected rows, extracts only the visible rows, sorts those rows in descending order by column L (index 11), writes the sorted data back to the original positions, and saves the file as SortedIgnoringHiddenRows.xlsx.
// Keywords: Aspose.Cells sort hidden rows | ignore hidden rows Aspose.Cells .NET | sort column descending Aspose.Cells | DataSorter visible rows only | C# Aspose.Cells sorting example | Excel hidden rows sort bypass
// Common Searches: Aspose.Cells sort column L descending ignoring hidden rows | How to skip hidden rows when sorting with Aspose.Cells | C# example for sorting visible rows only in Aspose.Cells | DataSorter hide rows Aspose.Cells .NET | Sort Excel data while preserving hidden rows using Aspose
// Developer Intent: Sort column L in descending order while excluding any hidden rows from the sort operation.
// Use Cases: Maintain subtotal or grouping rows hidden from users while sorting the displayed data. | Prepare a spreadsheet for export where hidden rows must stay in their original order. | Implement a web‑based reporting tool that sorts only visible rows in an Aspose.Cells workbook.
// AI Prompts: Show me a concise Aspose.Cells .NET code snippet that sorts column L descending and leaves hidden rows untouched. | How can I configure Aspose.Cells DataSorter to ignore hidden rows during a sort? | Provide a step‑by‑step explanation for sorting only visible rows in a workbook using C# and Aspose.Cells.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This example creates a workbook, fills columns A‑M with sample data, hides selected rows, extracts only the visible rows, sorts those rows in descending order by column L (index 11), writes the sorted data back to the original positions, and saves the file as SortedIgnoringHiddenRows.xlsx.
    class DataSorterIgnoreHiddenRowsDemo
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data (columns A to M, rows 1 to 10)
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col <= 12; col++) // column L is index 11
                {
                    cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
                // Put numeric values in column L for sorting
                cells[row, 11].PutValue(10 - row); // descending values initially
            }

            // Hide a few rows to demonstrate that they will be ignored during sorting
            cells.Rows[2].IsHidden = true; // hide row 3
            cells.Rows[5].IsHidden = true; // hide row 6

            // Determine the used range
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // Collect indices of visible rows
            List<int> visibleRowIndices = new List<int>();
            for (int r = 0; r <= maxRow; r++)
            {
                if (!cells.Rows[r].IsHidden)
                    visibleRowIndices.Add(r);
            }

            // Extract data of visible rows
            List<object[]> visibleRowsData = new List<object[]>();
            foreach (int r in visibleRowIndices)
            {
                object[] rowData = new object[maxCol + 1];
                for (int c = 0; c <= maxCol; c++)
                {
                    rowData[c] = cells[r, c].Value;
                }
                visibleRowsData.Add(rowData);
            }

            // Sort the extracted rows descending by column L (index 11)
            visibleRowsData.Sort((a, b) =>
            {
                // Handle possible nulls
                object valA = a[11];
                object valB = b[11];
                double numA = valA == null ? double.MinValue : Convert.ToDouble(valA);
                double numB = valB == null ? double.MinValue : Convert.ToDouble(valB);
                // Descending order
                return numB.CompareTo(numA);
            });

            // Write the sorted data back to the original visible rows
            for (int i = 0; i < visibleRowIndices.Count; i++)
            {
                int targetRow = visibleRowIndices[i];
                object[] rowData = visibleRowsData[i];
                for (int c = 0; c <= maxCol; c++)
                {
                    cells[targetRow, c].PutValue(rowData[c]);
                }
            }

            // Save the workbook
            workbook.Save("SortedIgnoringHiddenRows.xlsx");
        }
    }
}

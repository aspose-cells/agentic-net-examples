// Title: Select ListBox items in C# by matching values from column A of an Excel workbook using Aspose.Cells
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, iterates through each non‑empty cell in column A, and marks the matching entries as selected in a ListBox‑like collection. | Update the selection routine to perform case‑insensitive matching and correctly handle duplicate values when syncing Excel column A data with ListBox items. | Create a console utility that accepts an Excel file path as a command‑line argument, uses Aspose.Cells to read column A, and outputs the indices of ListBox items that match the spreadsheet values.
// Common Searches: c# aspose.cells read column a and select matching items in a listbox control | how to sync excel column values with listbox selections in a .net console application | select multiple items in a ListBox based on values from an xlsx file using Aspose.Cells
// Tags: Aspose.Cells read column A from XLSX | C# select ListBox items by matching Excel values | listbox selection based on spreadsheet data | console app Excel to ListBox synchronization | case-insensitive matching ListBox Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Minimal ListBox replacement for console applications
    // Loads an Excel workbook with Aspose.Cells, walks down column A until a blank cell, finds exact (or case‑insensitive) matches in a SimpleListBox, and marks those indices as selected.
    public class SimpleListBox
    {
        public List<string> Items { get; } = new List<string>();
        private readonly HashSet<int> _selectedIndices = new HashSet<int>();

        public const int NoMatches = -1;

        // Finds the exact string in Items and returns its index or NoMatches
        public int FindStringExact(string value)
        {
            for (int i = 0; i < Items.Count; i++)
            {
                if (string.Equals(Items[i], value, StringComparison.Ordinal))
                    return i;
            }
            return NoMatches;
        }

        // Marks the item at index as selected
        public void SetSelected(int index, bool selected)
        {
            if (selected)
                _selectedIndices.Add(index);
            else
                _selectedIndices.Remove(index);
        }

        // Clears all selections
        public void ClearSelected()
        {
            _selectedIndices.Clear();
        }

        // Returns the currently selected items
        public IEnumerable<string> GetSelectedItems()
        {
            foreach (int idx in _selectedIndices)
            {
                if (idx >= 0 && idx < Items.Count)
                    yield return Items[idx];
            }
        }
    }

    // Handles selection of ListBox items based on Excel column A values
    public class ListBoxSelector
    {
        private readonly SimpleListBox _listBox;
        private readonly string _excelFilePath;

        public ListBoxSelector(SimpleListBox listBox, string excelFilePath)
        {
            _listBox = listBox;
            _excelFilePath = excelFilePath;
        }

        public void SelectItemsFromColumnA()
        {
            try
            {
                // Verify that the Excel file exists
                if (!File.Exists(_excelFilePath))
                {
                    Console.WriteLine($"Error: File not found - {_excelFilePath}");
                    return;
                }

                // Load the workbook from the specified file
                Workbook workbook = new Workbook(_excelFilePath);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Clear any previous selections
                _listBox.ClearSelected();

                // Iterate through column A until a blank cell is encountered
                int row = 0;
                while (true)
                {
                    Cell cell = cells[row, 0];
                    // Stop when the cell is null, has a null value, or its string representation is empty/whitespace
                    if (cell == null ||
                        cell.Type == CellValueType.IsNull ||
                        string.IsNullOrWhiteSpace(cell.StringValue))
                    {
                        break;
                    }

                    string cellValue = cell.StringValue?.Trim();
                    if (!string.IsNullOrEmpty(cellValue))
                    {
                        int index = _listBox.FindStringExact(cellValue);
                        if (index != SimpleListBox.NoMatches)
                        {
                            _listBox.SetSelected(index, true);
                        }
                    }

                    row++;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception during selection: {ex.Message}");
            }
        }
    }

    // Console application entry point
    public static class Program
    {
        [STAThread]
        public static void Main()
        {
            // Initialize a simple ListBox with sample items
            var listBox = new SimpleListBox();
            listBox.Items.AddRange(new[] { "Apple", "Banana", "Cherry", "Date", "Elderberry" });

            // Path to the Excel file containing values in column A
            string excelPath = @"C:\Data\Values.xlsx";

            // Perform selection based on Excel data
            var selector = new ListBoxSelector(listBox, excelPath);
            selector.SelectItemsFromColumnA();

            // Output selected items
            Console.WriteLine("Selected items:");
            foreach (var item in listBox.GetSelectedItems())
            {
                Console.WriteLine($"- {item}");
            }
        }
    }
}

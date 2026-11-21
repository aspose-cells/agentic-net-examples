// Title: Read numeric cell values from an Excel workbook and store them in a List<double> using Aspose.Cells for .NET
// Common Searches: c# Aspose.Cells read only numeric cells from Excel and add to List<double> | how to collect numeric values from all worksheets with Aspose.Cells | Aspose.Cells iterate workbook cells and extract double values in .NET | efficiently retrieve numeric cell data from an Excel file using Aspose.Cells C# | store Excel numeric cell values in an in‑memory collection with Aspose.Cells
// Tags: numeric cell extraction Aspose.Cells .NET | collect double values from Excel workbook | in‑memory numeric collection C# | iterate worksheets cells Aspose.Cells | read numeric cells using CellValueType.IsNumeric | high‑performance Excel data extraction Aspose

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace LightCellsNumericReader
{
    // Collector that stores numeric cell values from a workbook
    // The example defines a NumericCellCollector that holds a List<double> of numeric values. Its Collect method loops through every worksheet and each cell in the workbook, checks for CellValueType.IsNumeric, and adds the cell's DoubleValue to the list. The program loads a workbook from "input.xlsx", uses the collector to gather numeric values, prints the total count, and outputs each value, with file‑existence validation and exception handling.
    public class NumericCellCollector
    {
        // In‑memory collection for numeric values
        public readonly List<double> NumericValues = new List<double>();

        // Collect numeric values from all worksheets in the given workbook
        public void Collect(Workbook workbook)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                // Iterate through all used cells
                foreach (Cell cell in cells)
                {
                    if (cell.Type == CellValueType.IsNumeric)
                    {
                        NumericValues.Add(cell.DoubleValue);
                    }
                }
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file
            string inputPath = "input.xlsx";

            try
            {
                // Ensure the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: File not found – {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Create the collector and gather numeric values
                NumericCellCollector collector = new NumericCellCollector();
                collector.Collect(workbook);

                // Output results
                Console.WriteLine($"Total numeric cells read: {collector.NumericValues.Count}");
                foreach (double val in collector.NumericValues)
                {
                    Console.WriteLine(val);
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsFindOldDates
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Retrieve the named range "DateRange"
                AsposeRange namedRange = workbook.Worksheets.GetRangeByName("DateRange");
                if (namedRange == null)
                {
                    Console.WriteLine("Named range 'DateRange' not found.");
                    return;
                }

                // Configure FindOptions (not used directly in the loop but kept for completeness)
                FindOptions findOptions = new FindOptions
                {
                    LookInType = LookInType.Values,
                    LookAtType = LookAtType.EntireContent,
                    SearchBackward = false,
                    SearchOrderByRows = true
                };

                // Define the search area based on the named range
                CellArea area = new CellArea
                {
                    StartRow = namedRange.FirstRow,
                    StartColumn = namedRange.FirstColumn,
                    EndRow = namedRange.FirstRow + namedRange.RowCount - 1,
                    EndColumn = namedRange.FirstColumn + namedRange.ColumnCount - 1
                };
                findOptions.SetRange(area);

                // Threshold date: one year ago from today
                DateTime thresholdDate = DateTime.Now.AddYears(-1);

                // Collect cells with dates older than the threshold
                List<Cell> oldDateCells = new List<Cell>();
                Worksheet sheet = namedRange.Worksheet;

                for (int row = area.StartRow; row <= area.EndRow; row++)
                {
                    for (int col = area.StartColumn; col <= area.EndColumn; col++)
                    {
                        Cell cell = sheet.Cells[row, col];
                        if (cell.Type == CellValueType.IsDateTime)
                        {
                            DateTime cellDate = cell.DateTimeValue;
                            if (cellDate < thresholdDate)
                            {
                                oldDateCells.Add(cell);
                            }
                        }
                    }
                }

                // Output results
                Console.WriteLine($"Cells with dates older than {thresholdDate:d}:");
                foreach (Cell c in oldDateCells)
                {
                    Console.WriteLine($"{c.Name} = {c.DateTimeValue:d}");
                }

                // Save the workbook (even if unchanged, to demonstrate successful processing)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
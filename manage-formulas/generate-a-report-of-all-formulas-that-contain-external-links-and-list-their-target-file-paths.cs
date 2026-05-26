using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsExternalLinkReport
{
    class Program
    {
        static void Main()
        {
            // Load the workbook (replace with your actual file path)
            string workbookPath = "input.xlsx";
            Workbook workbook = new Workbook(workbookPath);

            // List to hold report lines
            List<string> reportLines = new List<string>();

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Iterate through all cells in the worksheet
                foreach (Cell cell in cells)
                {
                    // Process only formula cells that contain external links
                    if (cell.IsFormula && cell.ContainsExternalLink)
                    {
                        // Get all precedents (references) used in the formula
                        ReferredAreaCollection precedents = cell.GetPrecedents();

                        if (precedents != null)
                        {
                            foreach (ReferredArea area in precedents)
                            {
                                // Check if the reference is an external link
                                if (area.IsExternalLink)
                                {
                                    // Build a report line with sheet name, cell address and external file path
                                    string line = $"Sheet: {sheet.Name}, Cell: {cell.Name}, External File: {area.ExternalFileName}";
                                    reportLines.Add(line);
                                }
                            }
                        }
                    }
                }
            }

            // Output the report to the console
            Console.WriteLine("External Link Report:");
            foreach (string line in reportLines)
            {
                Console.WriteLine(line);
            }

            // Optionally, save the report to a text file
            System.IO.File.WriteAllLines("ExternalLinkReport.txt", reportLines);
        }
    }
}
using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsFindAcrossWorksheets
{
    class Program
    {
        static void Main()
        {
            // Load the workbook (replace with your actual file path)
            Workbook workbook = new Workbook("InputData.xlsx");

            // Define the value to search for (can be string, int, double, DateTime, bool)
            string searchValue = "TargetValue";

            // Configure find options: search in cell values, match any part of the content
            FindOptions findOptions = new FindOptions
            {
                LookInType = LookInType.Values,
                LookAtType = LookAtType.Contains,
                SearchBackward = false,
                SeachOrderByRows = true
            };

            // List to hold all found cell addresses across worksheets
            List<string> foundCells = new List<string>();

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                Cell previousCell = null;

                // Repeatedly call Find to get all occurrences in the current sheet
                while (true)
                {
                    Cell found = cells.Find(searchValue, previousCell, findOptions);
                    if (found == null)
                        break; // No more matches in this sheet

                    // Record the worksheet name and cell address
                    foundCells.Add($"{sheet.Name}!{found.Name}");

                    // Set previousCell to the current found cell to continue searching
                    previousCell = found;
                }
            }

            // Output the results
            Console.WriteLine("Cells matching the criteria:");
            foreach (string address in foundCells)
            {
                Console.WriteLine(address);
            }

            // Optionally, you could highlight the found cells (example: set background color)
            // This part is optional and demonstrates further manipulation.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (string addr in foundCells)
                {
                    // addr format: SheetName!A1
                    string[] parts = addr.Split('!');
                    if (parts.Length != 2) continue;
                    if (parts[0] != sheet.Name) continue;

                    Cell cell = sheet.Cells[parts[1]];
                    Style style = cell.GetStyle();
                    style.ForegroundColor = System.Drawing.Color.Yellow;
                    style.Pattern = BackgroundType.Solid;
                    cell.SetStyle(style);
                }
            }

            // Save the workbook (replace with desired output path)
            workbook.Save("OutputData.xlsx");
        }
    }
}
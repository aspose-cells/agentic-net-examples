using System;
using Aspose.Cells;

namespace AsposeCellsFindDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("Input.xlsx");
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Search for a string value
            string searchString = "Target";
            Cell foundStringCell = cells.Find(searchString, null);
            if (foundStringCell != null)
            {
                Console.WriteLine($"String \"{searchString}\" found at: {foundStringCell.Name}");
                // Highlight the found cell
                Style style = workbook.CreateStyle();
                style.ForegroundColor = System.Drawing.Color.Yellow;
                style.Pattern = BackgroundType.Solid;
                foundStringCell.SetStyle(style);
            }
            else
            {
                Console.WriteLine($"String \"{searchString}\" not found.");
            }

            // Search for a numeric value using FindOptions (exact match)
            double searchNumber = 123.45;
            FindOptions numOptions = new FindOptions
            {
                LookInType = LookInType.Values,
                LookAtType = LookAtType.EntireContent,
                ValueTypeSensitive = true   // ensure numeric comparison
            };
            Cell foundNumberCell = cells.Find(searchNumber, null, numOptions);
            if (foundNumberCell != null)
            {
                Console.WriteLine($"Number {searchNumber} found at: {foundNumberCell.Name}");
                // Highlight the found cell with a different color
                Style numStyle = workbook.CreateStyle();
                numStyle.ForegroundColor = System.Drawing.Color.LightGreen;
                numStyle.Pattern = BackgroundType.Solid;
                foundNumberCell.SetStyle(numStyle);
            }
            else
            {
                Console.WriteLine($"Number {searchNumber} not found.");
            }

            // Save the workbook with highlights (replace with desired output path)
            workbook.Save("Output.xlsx");
        }
    }
}
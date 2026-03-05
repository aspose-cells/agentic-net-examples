using System;
using Aspose.Cells;

namespace AsposeCellsFindExample
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook
            Workbook workbook = new Workbook("InputData.xlsx");
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the value to search for
            string searchValue = "TargetText";

            // Configure find options (search in values, case‑insensitive, contains)
            FindOptions findOptions = new FindOptions
            {
                LookInType = LookInType.Values,
                LookAtType = LookAtType.Contains,
                CaseSensitive = false
            };

            // Perform the search starting from the beginning (previousCell = null)
            Cell foundCell = worksheet.Cells.Find(searchValue, null, findOptions);

            if (foundCell != null)
            {
                Console.WriteLine($"Found '{searchValue}' at cell {foundCell.Name} (Row {foundCell.Row}, Column {foundCell.Column}).");

                // Highlight the found cell with a red bold font
                Style highlightStyle = workbook.CreateStyle();
                highlightStyle.Font.Color = System.Drawing.Color.Red;
                highlightStyle.Font.IsBold = true;
                foundCell.SetStyle(highlightStyle);
            }
            else
            {
                Console.WriteLine($"Value '{searchValue}' was not found in the worksheet.");
            }

            // Save the workbook (the highlighted cell will be persisted)
            workbook.Save("OutputData.xlsx");
        }
    }
}
using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsFieldFormatting
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file
            string inputPath = "input.xlsx";

            // Load the workbook (create, load and save follow the standard lifecycle)
            Workbook workbook = new Workbook(inputPath);

            // Regular expression to match the {{field_name:FORMAT}} pattern
            Regex fieldPattern = new Regex(@"\{\{([^:}]+):([^}]+)\}\}", RegexOptions.Compiled);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Iterate through all used cells
                foreach (Cell cell in cells)
                {
                    // Process only string cells
                    if (cell.Type == CellValueType.IsString)
                    {
                        string cellText = cell.StringValue;
                        Match match = fieldPattern.Match(cellText);
                        if (match.Success && match.Groups.Count == 3)
                        {
                            // Extract field name and format string
                            string fieldName = match.Groups[1].Value;
                            string formatString = match.Groups[2].Value;

                            // Replace the cell value with the field name (or keep original as needed)
                            cell.PutValue(fieldName);

                            // Create a new style with the custom format
                            Style style = workbook.CreateStyle();
                            style.Custom = formatString; // Apply the exact format

                            // Assign the style to the cell
                            cell.SetStyle(style);
                        }
                    }
                }
            }

            // Save the modified workbook
            string outputPath = "output_formatted.xlsx";
            workbook.Save(outputPath);
        }
    }
}
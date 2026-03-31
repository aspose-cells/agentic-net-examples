using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsDateFormatExample
{
    class Program
    {
        static void Main(string[] args)
        {
            ApplyDateFormat.Run();
        }
    }

    public class ApplyDateFormat
    {
        public static void Run()
        {
            // Path to the source XLSX file
            string inputPath = "input.xlsx";

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the cells that should be formatted as dates
            // Example: A2, B5, C10 – replace with your actual cell addresses
            List<string> dateCellAddresses = new List<string>
            {
                "A2",
                "B5",
                "C10"
            };

            // Iterate over each designated cell and apply the MM/dd/yyyy format
            foreach (string address in dateCellAddresses)
            {
                // Access the cell
                Cell cell = worksheet.Cells[address];

                // Ensure the cell contains a DateTime value; if it's a string, try to parse it
                if (cell.Type == CellValueType.IsString)
                {
                    if (DateTime.TryParse(cell.StringValue, out DateTime parsedDate))
                    {
                        cell.PutValue(parsedDate);
                    }
                }

                // Retrieve the current style of the cell
                Style style = cell.GetStyle();

                // Apply a custom date format (MM/dd/yyyy)
                style.Custom = "MM/dd/yyyy";

                // Assign the modified style back to the cell
                cell.SetStyle(style);
            }

            // Save the modified workbook to a new file
            string outputPath = "output_formatted.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook saved with date formatting applied to cells: {string.Join(", ", dateCellAddresses)}");
        }
    }
}
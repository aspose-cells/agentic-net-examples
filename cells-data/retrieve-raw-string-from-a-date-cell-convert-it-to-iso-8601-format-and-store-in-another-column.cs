// Title: C# – Convert Excel date strings to ISO‑8601 using Aspose.Cells
// Description: Loads an Excel workbook, reads the raw text from each used cell in column A, parses it to a DateTime, formats the value as ISO‑8601 (yyyy‑MM‑ddTHH:mm:ss), writes the formatted string to column B, and saves the result to a new file. Includes handling for missing files, unparsable dates, and uses MaxDataRow to limit the iteration to populated rows.
// Keywords: Aspose.Cells C# date conversion | Excel raw string to ISO 8601 | parse text date Aspose.Cells | write formatted date column B | MaxDataRow last row Aspose | C# Excel ISO 8601 example
// Common Searches: Aspose.Cells convert text date to ISO 8601 C# | read raw date string from Excel cell using Aspose | write ISO formatted date to adjacent column in .NET | determine last used row Aspose.Cells MaxDataRow | handle unparsable dates Aspose.Cells C#
// Developer Intent: Read date strings from column A, convert each to ISO‑8601, and store the result in column B of the same worksheet.
// Use Cases: Export data to APIs that require ISO‑8601 timestamps. | Normalize legacy spreadsheets that store dates as plain text. | Prepare workbooks for downstream systems that enforce a uniform date format. | Generate audit‑ready reports with consistent timestamp representation.
// AI Prompts: Create a C# Aspose.Cells snippet that reads a text date from column A, parses it, formats it as "yyyy-MM-ddTHH:mm:ss", writes the result to column B, and skips rows with invalid dates. | Show how to use MaxDataRow to iterate only the populated rows when converting dates in an Excel sheet with Aspose.Cells. | Explain error handling for missing input files and date‑parsing failures in an Aspose.Cells date‑conversion routine.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDateConversion
{
    // Loads an Excel workbook, reads the raw text from each used cell in column A, parses it to a DateTime, formats the value as ISO‑8601 (yyyy‑MM‑ddTHH:mm:ss), writes the formatted string to column B, and saves the result to a new file. Includes handling for missing files, unparsable dates, and uses MaxDataRow to limit the iteration to populated rows.
    public class ConvertDateToIso
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Path to the source workbook (replace with actual file path)
            string inputPath = "input.xlsx";

            // Path to the destination workbook
            string outputPath = "output.xlsx";

            // Ensure the input file exists; create a placeholder if it does not
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}. Creating a new workbook.");
                Workbook placeholder = new Workbook();
                placeholder.Worksheets[0].Name = "Sheet1";
                placeholder.Save(inputPath);
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Determine the last row that contains data in column A (zero‑based index)
                int lastRow = cells.MaxDataRow;

                // Iterate through each used row in column A
                for (int row = 0; row <= lastRow; row++)
                {
                    // Get the cell that contains the original date string (Column A)
                    Cell sourceCell = cells[row, 0];

                    // Retrieve the raw string as it appears in the cell
                    string rawDateString = sourceCell.StringValue;

                    // Try to parse the string into a DateTime object
                    if (DateTime.TryParse(rawDateString, out DateTime parsedDate))
                    {
                        // Convert the DateTime to ISO 8601 format (e.g., 2023-05-15T00:00:00)
                        string isoDateString = parsedDate.ToString("yyyy-MM-ddTHH:mm:ss");

                        // Store the ISO string in the adjacent column (Column B)
                        Cell targetCell = cells[row, 1];
                        targetCell.PutValue(isoDateString);
                    }
                    else
                    {
                        // Write an empty string if parsing fails to keep columns aligned
                        cells[row, 1].PutValue(string.Empty);
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}.");
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.WriteLine($"File not found: {fnfEx.FileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}

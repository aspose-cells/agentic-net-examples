using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsLineBreakValidation
{
    public class ValidateCellLineBreaks
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle: create)
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Define a string with explicit line breaks
                string originalText = "First line\nSecond line\r\nThird line";

                // Put the text into cell A1
                Cell cell = sheet.Cells["A1"];
                cell.PutValue(originalText);

                // Enable text wrapping so line breaks are respected in the UI
                Style style = cell.GetStyle();
                style.IsTextWrapped = true;
                cell.SetStyle(style);

                // Auto‑fit the row height to display all lines without extra spacing
                sheet.AutoFitRow(0);

                // Save the workbook (lifecycle: save)
                string filePath = "LineBreakValidation.xlsx";
                workbook.Save(filePath);

                // Verify the file exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"Error: File '{filePath}' was not created.");
                    return;
                }

                // Reload the workbook (lifecycle: load) to verify the stored value
                Workbook loadedWorkbook = new Workbook(filePath);
                Cell loadedCell = loadedWorkbook.Worksheets[0].Cells["A1"];
                string loadedText = loadedCell.StringValue;

                // Validation: the loaded text should match the original text exactly
                bool isValid = string.Equals(originalText, loadedText, StringComparison.Ordinal);
                Console.WriteLine(isValid
                    ? "Validation succeeded: line breaks are preserved without extra spacing."
                    : "Validation failed: line break content differs.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            ValidateCellLineBreaks.Run();
        }
    }
}
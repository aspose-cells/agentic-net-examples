using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class SetCharactersWithRangeCheck
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cell cell = worksheet.Cells["A1"];

                // Set the cell value that will be formatted
                cell.Value = "Hello Aspose.Cells!";

                // Prepare font settings for different parts of the text
                FontSetting[] fontSettings = new FontSetting[2];

                // First part: characters 0-4 ("Hello")
                fontSettings[0] = cell.Characters(0, 5);
                fontSettings[0].Font.IsBold = true;
                fontSettings[0].Font.Color = Color.Red;

                // Second part: characters 6-15 ("Aspose.Cel")
                // Intentionally use a length that may exceed the text length to demonstrate error handling
                int startIndex = 6;
                int length = 20; // This exceeds the actual remaining characters
                fontSettings[1] = cell.Characters(startIndex, length);
                fontSettings[1].Font.IsItalic = true;
                fontSettings[1].Font.Color = Color.Blue;

                // Validate each FontSetting range against the actual cell text length
                string cellText = cell.StringValue ?? string.Empty;
                int textLength = cellText.Length;

                foreach (var fs in fontSettings)
                {
                    int endIndex = fs.StartIndex + fs.Length;
                    if (endIndex > textLength)
                    {
                        throw new ArgumentOutOfRangeException(
                            $"The character range (StartIndex={fs.StartIndex}, Length={fs.Length}) exceeds the cell text length ({textLength}).");
                    }
                }

                // Apply the rich text formatting
                cell.SetCharacters(fontSettings);
                Console.WriteLine("SetCharacters executed successfully.");

                // Save the workbook
                workbook.Save("SetCharactersWithRangeCheck.xlsx");
            }
            catch (ArgumentOutOfRangeException ex)
            {
                // Handle range errors specifically
                Console.WriteLine($"Range error: {ex.Message}");
            }
            catch (Exception ex)
            {
                // Handle any other unexpected errors
                Console.WriteLine($"Error executing SetCharacters: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SetCharactersWithRangeCheck.Run();
        }
    }
}
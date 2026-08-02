using System;
using System.Drawing;
using Aspose.Cells;

class SetCharactersErrorHandlingDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cell cell = worksheet.Cells["A1"];

        // Set initial text in the cell
        cell.PutValue("Hello World");

        // Define formatting ranges (some may exceed the cell text length)
        var ranges = new (int start, int length, Color color)[]
        {
            (0, 5, Color.Red),      // Valid range: "Hello"
            (6, 10, Color.Blue)     // Exceeds length: start 6, length 10 (cell length is 11)
        };

        // Retrieve the cell text length for validation
        string cellText = cell.StringValue ?? string.Empty;
        int textLength = cellText.Length;

        // Prepare an array to hold FontSetting objects
        FontSetting[] fontSettings = new FontSetting[ranges.Length];
        int validCount = 0;

        // Iterate over each range, validate, and create FontSetting objects
        for (int i = 0; i < ranges.Length; i++)
        {
            int start = ranges[i].start;
            int length = ranges[i].length;

            // Validate start index
            if (start < 0 || start > textLength)
            {
                Console.WriteLine($"Start index {start} is out of bounds for text length {textLength}. Skipping this range.");
                continue;
            }

            // Adjust length if it exceeds the remaining characters
            if (start + length > textLength)
            {
                Console.WriteLine($"Requested range ({start}, {length}) exceeds text length {textLength}. Adjusting length to fit.");
                length = textLength - start;
            }

            // Create the Characters object and apply the desired color
            FontSetting setting = cell.Characters(start, length);
            setting.Font.Color = ranges[i].color;

            // Store the valid setting
            fontSettings[validCount++] = setting;
        }

        // Trim the array to contain only the valid settings
        Array.Resize(ref fontSettings, validCount);

        // Apply the rich text formatting with error handling
        try
        {
            cell.SetCharacters(fontSettings);
            Console.WriteLine("SetCharacters executed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error executing SetCharacters: {ex.Message}");
        }

        // Save the workbook
        workbook.Save("SetCharactersErrorHandlingDemo.xlsx");
    }
}
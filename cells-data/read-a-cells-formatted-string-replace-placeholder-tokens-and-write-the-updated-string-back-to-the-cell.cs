using System;
using Aspose.Cells;

namespace AsposeCellsPlaceholderReplacement
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Put a sample text containing placeholders into cell A1
            // Example placeholder format: {Name}
            cells["A1"].PutValue("Dear {Name}, your order {OrderId} is confirmed.");

            // Define replace options (case‑insensitive, replace within the cell content)
            ReplaceOptions options = new ReplaceOptions
            {
                CaseSensitive = false,
                MatchEntireCellContents = false
            };

            // Replace the placeholders with actual values
            cells["A1"].Replace("{Name}", "John Doe", options);
            cells["A1"].Replace("{OrderId}", "12345", options);

            // Optionally, read the formatted string after replacement
            string result = cells["A1"].StringValue;
            Console.WriteLine("Updated cell value: " + result);

            // Save the workbook
            workbook.Save("PlaceholderReplacement.xlsx", SaveFormat.Xlsx);
        }
    }
}
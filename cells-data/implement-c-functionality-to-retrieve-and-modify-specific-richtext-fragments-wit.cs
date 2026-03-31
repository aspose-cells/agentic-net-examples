using System;
using System.Drawing;
using Aspose.Cells;

namespace RichTextFragmentDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Get a reference to cell A1 and set an initial plain text value
            Cell cell = worksheet.Cells["A1"];
            cell.PutValue("Hello World");

            // Check if the cell currently contains rich‑text formatting
            if (cell.IsRichText())
            {
                // Retrieve existing rich‑text fragments (if any)
                FontSetting[] existingFragments = cell.GetCharacters();
                Console.WriteLine($"Existing rich‑text fragments count: {existingFragments.Length}");
            }

            // Define rich‑text formatting for the first part "Hello"
            FontSetting helloFragment = cell.Characters(0, 5); // start index, length
            helloFragment.Font.IsBold = true;
            helloFragment.Font.Color = Color.Red;

            // Define rich‑text formatting for the second part "World"
            FontSetting worldFragment = cell.Characters(6, 5); // start index, length
            worldFragment.Font.IsItalic = true;
            worldFragment.Font.Color = Color.Blue;

            // Apply the formatted fragments back to the cell
            cell.SetCharacters(new FontSetting[] { helloFragment, worldFragment });

            // Optionally retrieve the rich value object (demonstration of GetRichValue)
            CellRichValue richValue = cell.GetRichValue();
            Console.WriteLine($"Rich value type: {richValue?.GetType().Name ?? "null"}");

            // Save the workbook to disk (lifecycle: save)
            workbook.Save("RichTextModified.xlsx");
        }
    }
}
// Title: C# – Serialize Aspose.Cells RichTextPortion (FontSetting) Collection to JSON
// Description: Demonstrates how to create a workbook, apply bold red and italic blue formatting to parts of cell A1, extract the FontSetting array with GetCharacters(), build a lightweight object containing start index, length and font attributes (bold, italic, underline, ARGB color), and serialize the collection to an indented JSON string that can be saved and reused.
// Keywords: Aspose.Cells | C# | RichTextPortion serialization | FontSetting to JSON | export cell formatting | JSON rich text | save Aspose.Cells styles | deserialize FontSetting
// Common Searches: serialize Aspose.Cells rich text to JSON C# | export FontSetting collection as JSON | save cell rich text formatting Aspose.Cells | how to get characters from a cell Aspose.Cells | store Aspose.Cells cell styles in JSON
// Developer Intent: Convert a cell's RichTextPortion (FontSetting) collection into a JSON representation for later restoration or external processing.
// Use Cases: Persist custom cell formatting in a configuration file and reapply it to other workbooks. | Exchange rich‑text styling between services by transmitting JSON metadata. | Archive cell style details in a database for reporting, auditing, or version control.
// AI Prompts: Write C# code that reads the generated RichTextPortions.json, deserializes the objects, and reapplies the formatting to a target cell using Aspose.Cells. | Provide a helper method that converts the stored ARGB integer back to System.Drawing.Color when restoring font colors from JSON. | Explain how to safely handle missing underline information during deserialization of the rich‑text JSON.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// Demonstrates how to create a workbook, apply bold red and italic blue formatting to parts of cell A1, extract the FontSetting array with GetCharacters(), build a lightweight object containing start index, length and font attributes (bold, italic, underline, ARGB color), and serialize the collection to an indented JSON string that can be saved and reused.
class SerializeRichTextPortions
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Access cell A1 and set a plain text value
        Cell cell = worksheet.Cells["A1"];
        cell.PutValue("Hello World!");

        // Apply rich text formatting to different parts of the cell text
        // Format "Hello" as bold red
        FontSetting helloPortion = cell.Characters(0, 5);
        helloPortion.Font.IsBold = true;
        helloPortion.Font.Color = Color.Red;

        // Format "World" as italic blue
        FontSetting worldPortion = cell.Characters(6, 5);
        worldPortion.Font.IsItalic = true;
        worldPortion.Font.Color = Color.Blue;

        // Verify that the cell contains rich text
        if (!cell.IsRichText())
        {
            Console.WriteLine("The cell does not contain rich text.");
            return;
        }

        // Retrieve all rich text portions (FontSetting objects) from the cell
        FontSetting[] richPortions = cell.GetCharacters();

        // Convert the FontSetting collection into a serializable structure
        var serializableList = new List<object>();
        foreach (FontSetting fs in richPortions)
        {
            var portionInfo = new
            {
                StartIndex = fs.StartIndex,
                Length = fs.Length,
                Font = new
                {
                    IsBold = fs.Font.IsBold,
                    IsItalic = fs.Font.IsItalic,
                    IsUnderline = fs.Font.Underline != FontUnderlineType.None,
                    // Store color as ARGB integer for simplicity
                    ColorArgb = fs.Font.Color.ToArgb()
                }
            };
            serializableList.Add(portionInfo);
        }

        // Serialize the list to a formatted JSON string
        string json = JsonSerializer.Serialize(serializableList, new JsonSerializerOptions { WriteIndented = true });

        // Output the JSON to console
        Console.WriteLine("RichTextPortion collection serialized to JSON:");
        Console.WriteLine(json);

        // Optionally, save the JSON to a file for later reuse
        File.WriteAllText("RichTextPortions.json", json);
    }
}

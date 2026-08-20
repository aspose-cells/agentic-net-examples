// Title: Preserve Rich‑Text Formatting When Saving and Reloading an Aspose.Cells Workbook (C#)
// AI Prompts: Show how to apply bold formatting to a specific character range in a cell and verify it persists after saving to XLSX with Aspose.Cells. | Provide C# code that saves a workbook, reloads it, and checks that character‑level formatting (e.g., bold) remains unchanged.
// Common Searches: Aspose.Cells keep rich text formatting after save C# | verify character formatting after reloading workbook Aspose.Cells | preserve bold text in cell A1 when saving to XLSX Aspose.Cells | IsRichText method usage Aspose.Cells .NET | save workbook with partial bold text Aspose.Cells
// Tags: Aspose.Cells | C# | Rich Text | Character Formatting | SaveFormat.Xlsx | Preserve Formatting | IsRichText | Workbook Reload

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsRichTextPreserveDemo
{
    // Demonstrates creating a workbook, applying bold formatting to part of a cell's text, saving to XLSX, reloading the file, and confirming that the rich‑text attributes (bold) are retained.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Target cell for rich‑text
                Cell cell = worksheet.Cells["A1"];
                cell.PutValue("Hello World");

                // Apply rich‑text formatting:
                // Make "Hello" bold, keep the rest normal
                // Characters are zero‑based; length of "Hello" is 5
                cell.Characters(0, 5).Font.IsBold = true;
                cell.Characters(5, cell.StringValue.Length - 5).Font.IsBold = false;

                // Verify that the cell now contains rich text
                bool isRich = cell.IsRichText(); // IsRichText is a method in Aspose.Cells
                Console.WriteLine($"Cell A1 is rich text: {isRich}");

                // Save the workbook (preserves all formatting, including rich text)
                string filePath = "RichTextPreserved.xlsx";
                workbook.Save(filePath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {filePath}");

                // Reload the workbook to confirm formatting persisted
                if (File.Exists(filePath))
                {
                    try
                    {
                        Workbook loadedWorkbook = new Workbook(filePath);
                        Cell loadedCell = loadedWorkbook.Worksheets[0].Cells["A1"];

                        // Check formatting of the first character range
                        bool firstPartBold = loadedCell.Characters(0, 5).Font.IsBold;
                        bool secondPartBold = loadedCell.Characters(5, loadedCell.StringValue.Length - 5).Font.IsBold;

                        Console.WriteLine($"After reload - first part bold: {firstPartBold}");
                        Console.WriteLine($"After reload - second part bold: {secondPartBold}");
                    }
                    catch (Exception loadEx)
                    {
                        Console.WriteLine($"Error loading workbook: {loadEx.Message}");
                    }
                }
                else
                {
                    Console.WriteLine($"Error: File '{filePath}' was not found after saving.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

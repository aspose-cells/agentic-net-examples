// Title: C# – Retrieve and Format a Rich‑Text Portion with Cell.GetCharacters in Aspose.Cells
// Description: This example creates a workbook, writes "HelloWorld" to cell A1, calls Cell.GetCharacters(5,5) to obtain a FontSetting for the substring "World", applies bold and blue styling, and saves the file, demonstrating partial text formatting in Excel via Aspose.Cells for .NET.
// Keywords: Aspose.Cells | Cell.GetCharacters | .NET | C# Excel | rich text portion | partial cell formatting | FontSetting | modify cell characters | Excel SDK | Aspose.Cells example
// Common Searches: Cell.GetCharacters Aspose.Cells C# | format part of cell text Aspose | retrieve characters from Excel cell .NET | partial string formatting in Aspose.Cells | how to bold specific characters in Excel using Aspose
// Developer Intent: Extract a substring from a cell and change its font attributes.
// Use Cases: Emphasize a keyword inside a status message | Color‑code sections of a product code | Apply conditional styling to overdue dates within a single cell | Create multi‑style headers in a report | Highlight error codes embedded in a description
// AI Prompts: Write C# code using Aspose.Cells to make characters 2‑6 of cell C3 italic and red. | Show how to loop through words in a cell and assign each a different font size with Cell.GetCharacters. | Explain combining Cell.GetCharacters with if‑else to format numeric substrings based on thresholds.

using System;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsExamples
{
    // This example creates a workbook, writes "HelloWorld" to cell A1, calls Cell.GetCharacters(5,5) to obtain a FontSetting for the substring "World", applies bold and blue styling, and saves the file, demonstrating partial text formatting in Excel via Aspose.Cells for .NET.
    public class RetrieveRichTextPortionDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Put a string value into cell A1
                Cell cell = worksheet.Cells["A1"];
                cell.PutValue("HelloWorld");

                // Retrieve a rich text portion (characters from index 5, length 5)
                // This returns a FontSetting object representing the specified range
                FontSetting richPortion = cell.Characters(5, 5);

                // Modify the retrieved portion's font
                richPortion.Font.IsBold = true;
                richPortion.Font.Color = Color.Blue;

                // Save the workbook to verify the changes
                workbook.Save("RichTextPortionDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            RetrieveRichTextPortionDemo.Run();
        }
    }
}

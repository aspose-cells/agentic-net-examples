// Title: Import a 2D string array and style the header row with Aspose.Cells (C#)
// Description: Creates a new Workbook, imports a two‑dimensional string array at cell A1, defines a bold blue 12‑pt font style, applies it to the first row via a Range, and saves the file as HeaderStyled.xlsx.
// Keywords: Aspose.Cells | ImportTwoDimensionArray | C# Excel export | header row style | custom font Aspose | range.ApplyStyle | Workbook.Save | Excel styling C#
// Common Searches: Aspose.Cells import 2D array C# | How to style header row in Aspose.Cells | Apply bold blue font to first row using Aspose.Cells | Create style and apply to range Aspose.Cells C# | Save workbook after formatting Aspose.Cells
// Developer Intent: Load a 2‑D string array into a worksheet and highlight the header row with a custom font.
// Use Cases: Export a product catalog where column titles appear in bold blue for quick identification. | Generate a sales report that imports raw data and emphasizes the header for better readability. | Build an automated Excel export feature that formats the first row after loading data arrays.
// AI Prompts: Write C# code using Aspose.Cells to import a 2D string array and apply a red italic font to the header row. | Explain how to create a Style object, set font properties, and apply it to a specific range after using ImportTwoDimensionArray. | Show how to change the header font size and color dynamically based on user input in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Creates a new Workbook, imports a two‑dimensional string array at cell A1, defines a bold blue 12‑pt font style, applies it to the first row via a Range, and saves the file as HeaderStyled.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Two‑dimensional array of strings (header + data)
            string[,] data = new string[,]
            {
                { "Product", "Price", "Quantity" },
                { "Apple",  "1.50",  "10" },
                { "Banana", "0.75",  "20" },
                { "Orange", "1.20",  "15" }
            };

            // Import the array starting at cell A1 (row 0, column 0)
            cells.ImportTwoDimensionArray(data, 0, 0);

            // Create a style for the header row
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.Font.Color = Color.Blue;
            headerStyle.Font.Size = 12;

            // Apply the style to the first row (header)
            int columnCount = data.GetLength(1);
            AsposeRange headerRange = cells.CreateRange(0, 0, 1, columnCount);
            headerRange.ApplyStyle(headerStyle, null);

            // Save the workbook
            string outputPath = "HeaderStyled.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

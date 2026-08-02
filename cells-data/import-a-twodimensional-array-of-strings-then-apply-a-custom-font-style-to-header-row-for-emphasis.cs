// Title: Import a 2‑D string array and style the header row with Aspose.Cells for .NET
// Description: Creates a Workbook, imports a two‑dimensional string array of product data into the first worksheet at A1, builds a bold, blue, 12‑pt centered style, applies it to the header row via a Range and StyleFlag, auto‑fits columns, and saves the file as TwoDimArrayWithHeaderStyle.xlsx.
// Keywords: Aspose.Cells import 2D array C# | Apply header style Aspose.Cells | Bold blue font Excel .NET | StyleFlag all attributes | AutoFitColumns after import | CreateStyle Aspose.Cells | Range.ApplyStyle example
// Common Searches: how to import a 2d string array into Aspose.Cells worksheet | apply bold blue centered font to header row using Aspose.Cells | Aspose.Cells C# style header after ImportTwoDimensionArray | auto fit columns after importing data with Aspose.Cells | sample code for styling first row in Excel with Aspose.Cells
// Developer Intent: Load a 2‑D string array into an Excel sheet and format the first row with a custom font style.
// Use Cases: Generate a product catalog from in‑memory data and highlight column titles. | Create a sales summary where the header row needs emphasis before distribution. | Build quick Excel reports from runtime collections with styled headings for readability.
// AI Prompts: Write C# code using Aspose.Cells to import a two‑dimensional string array at A1 and apply a bold, blue, 12‑pt centered style to the first row. | Show how to use StyleFlag.All to apply a complete style to a header range and then auto‑fit the columns. | Explain how to extend the header style with background color and borders while keeping the existing font settings.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsTwoDimArrayDemo
{
    // Creates a Workbook, imports a two‑dimensional string array of product data into the first worksheet at A1, builds a bold, blue, 12‑pt centered style, applies it to the header row via a Range and StyleFlag, auto‑fits columns, and saves the file as TwoDimArrayWithHeaderStyle.xlsx.
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

                // Define a two‑dimensional array of strings (including header row)
                object[,] data = new object[,]
                {
                    { "Product", "Price", "Quantity" },   // Header row
                    { "Apple", "1.20", "50" },
                    { "Banana", "0.80", "30" },
                    { "Orange", "1.00", "40" }
                };

                // Import the array into the worksheet starting at cell A1 (row 0, column 0)
                cells.ImportTwoDimensionArray(data, 0, 0);

                // Create a custom style for the header row
                Style headerStyle = workbook.CreateStyle();
                headerStyle.Font.IsBold = true;                 // Bold font
                headerStyle.Font.Color = Color.Blue;            // Blue color
                headerStyle.Font.Size = 12;                     // Slightly larger font
                headerStyle.HorizontalAlignment = TextAlignmentType.Center; // Center align

                // Apply the style to the first row (header)
                int headerColumns = data.GetLength(1);
                Aspose.Cells.Range headerRange = cells.CreateRange(0, 0, 1, headerColumns);

                // Define which style attributes to apply (all in this case)
                StyleFlag flag = new StyleFlag { All = true };
                headerRange.ApplyStyle(headerStyle, flag);

                // Auto‑fit columns for better visibility
                worksheet.AutoFitColumns();

                // Save the workbook
                workbook.Save("TwoDimArrayWithHeaderStyle.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

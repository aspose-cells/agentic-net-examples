// Title: Insert an Image and Freeze Rows Above It with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to add a picture from a local file to a worksheet, then lock the rows above the graphic using the FreezePanes method, and finally save the workbook as an XLSX file.
// Keywords: Aspose.Cells | C# | insert picture | freeze panes | freeze rows | worksheet image | Excel header logo | FreezePanes method | add picture to Excel | Aspose.Cells for .NET
// Common Searches: Aspose.Cells add image to worksheet C# | How to freeze rows above a picture in Aspose.Cells | C# FreezePanes after inserting a logo with Aspose.Cells | Insert picture and keep it visible while scrolling Excel using Aspose | Aspose.Cells example: picture + frozen header rows
// Developer Intent: Place a graphic in a sheet and keep the upper rows static during scrolling.
// Use Cases: Create a report template with a company logo that remains on screen as users scroll through data. | Generate invoices where a seal or watermark is anchored at the top and the surrounding rows are frozen. | Build dashboards that display a banner image while allowing free navigation of the detail section.
// AI Prompts: Provide C# code that inserts a PNG at row 3 column B with Aspose.Cells and freezes the first three rows. | Show how to add multiple pictures at different rows and apply FreezePanes so all rows above the lowest image stay fixed. | Write an Aspose.Cells snippet that checks for image existence, inserts it, and then locks the rows above the inserted picture.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add a picture from a local file to a worksheet, then lock the rows above the graphic using the FreezePanes method, and finally save the workbook as an XLSX file.
class InsertImageAndFreezeRows
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the position where the image will be placed (row 5, column A)
            int topRow = 5;      // zero‑based index, so this is the 6th row in Excel
            int leftColumn = 0;  // column A

            // Resolve image path relative to the executable folder
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "image.jpg");

            // Add the picture if the file exists
            if (File.Exists(imagePath))
            {
                worksheet.Pictures.Add(topRow, leftColumn, imagePath);
            }
            else
            {
                Console.WriteLine($"Warning: Image file not found at '{imagePath}'. Skipping image insertion.");
            }

            // Freeze rows above the image (rows 0‑4) and no columns
            worksheet.FreezePanes(topRow, leftColumn, topRow, 0);

            // Save the workbook
            string outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "output.xlsx");
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

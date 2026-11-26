// Title: Insert a picture with alternative text into an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: how to embed an image with alt text in an Excel file using Aspose.Cells C# | Aspose.Cells C# add picture to specific cells and set accessibility description | set alternative text for a picture shape in an Aspose.Cells workbook | insert picture from file stream into worksheet and define alt text with Aspose.Cells | adjust picture dimensions and provide screen reader description in Excel using Aspose.Cells .NET
// Tags: add picture to worksheet Aspose.Cells C# | set picture alternative text Aspose.Cells | embed image from file stream Aspose.Cells | define picture size Excel .NET | accessibility alt text for Excel images Aspose.Cells | place picture across cell range Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, loads an image file via FileStream, inserts it as a picture spanning cells B2‑E6, assigns alternative (alt) text for screen‑reader accessibility, optionally sets width and height, and saves the file as PictureWithAltText.xlsx.
class Program
{
    static void Main(string[] args)
    {
        // ------------------------------------------------------------
        // 1. Determine the image file to embed.
        // ------------------------------------------------------------
        // If a command‑line argument is supplied, use it; otherwise fall back
        // to a file named "sample.jpg" in the current directory.
        string imagePath = args.Length > 0 ? args[0] : "sample.jpg";

        if (!File.Exists(imagePath))
        {
            Console.Error.WriteLine($"Error: Image file not found – \"{imagePath}\"");
            return;
        }

        // ------------------------------------------------------------
        // 2. Create a new workbook and obtain the first worksheet.
        // ------------------------------------------------------------
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // ------------------------------------------------------------
        // 3. Add the picture to the worksheet.
        // ------------------------------------------------------------
        // Position: start at cell B2 (row index 1, column index 1) and span
        // 5 rows × 5 columns. Adjust these indices as needed.
        int topRow = 1;          // B2 row (zero‑based)
        int leftColumn = 1;      // B2 column (zero‑based)
        int bottomRow = topRow + 4;
        int rightColumn = leftColumn + 4;

        Picture picture;
        using (FileStream fs = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
        {
            picture = worksheet.Shapes.AddPicture(topRow, leftColumn, bottomRow, rightColumn, fs);
        }

        // ------------------------------------------------------------
        // 4. Provide alternative (screen‑reader) text.
        // ------------------------------------------------------------
        // This text describes the visual content of the picture for users
        // relying on assistive technologies.
        picture.AlternativeText = "Company logo: a blue circle with white text reading \"Acme Corp\"";

        // Optional: fine‑tune the displayed size (pixels).
        picture.Width = 200;
        picture.Height = 100;

        // ------------------------------------------------------------
        // 5. Save the workbook.
        // ------------------------------------------------------------
        string outputFile = "PictureWithAltText.xlsx";
        try
        {
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook successfully saved to \"{outputFile}\"");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}

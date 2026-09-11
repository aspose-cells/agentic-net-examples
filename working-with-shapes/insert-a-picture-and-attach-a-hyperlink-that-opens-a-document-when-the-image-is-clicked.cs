// Title: Insert a PNG image into an Excel worksheet and attach a clickable hyperlink to a DOCX file using Aspose.Cells for .NET (C#)
// AI Prompts: Add a picture to cell B2 of a new workbook and set its Hyperlink.Address to a .docx file with Aspose.Cells in C#. | Create an Excel file, embed a PNG, and make the image open a target document when clicked using the Aspose.Cells API. | Programmatically attach a hyperlink to a Picture object in Aspose.Cells and save the workbook as an .xlsx file.
// Common Searches: Aspose.Cells C# insert image and link it to a Word document | set hyperlink on picture in Excel using Aspose.Cells for .NET | clickable image in generated XLSX that opens a DOCX file with Aspose.Cells
// Tags: insert picture Aspose.Cells C# | picture hyperlink Aspose.Cells | add png to worksheet Aspose.Cells | hyperlink property for picture Aspose.Cells | link image to docx Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, inserts a PNG image at cell B2, assigns a hyperlink that points to a DOCX file to the picture, and saves the result as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Path to the image that will be inserted
            string imagePath = "sample.png";

            // Ensure the image file exists to avoid FileNotFoundException
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            // Insert the picture at cell B2 (row index 1, column index 1)
            // (lifecycle rule: insert picture)
            int pictureIdx = sheet.Pictures.Add(1, 1, imagePath);
            Picture picture = sheet.Pictures[pictureIdx];

            // Attach a hyperlink to the picture that opens a document when clicked
            // The Hyperlink object is read‑only; set its Address property directly
            picture.Hyperlink.Address = "targetDocument.docx";

            // Save the workbook (lifecycle rule: save)
            string resultPath = "Result.xlsx";
            workbook.Save(resultPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {resultPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

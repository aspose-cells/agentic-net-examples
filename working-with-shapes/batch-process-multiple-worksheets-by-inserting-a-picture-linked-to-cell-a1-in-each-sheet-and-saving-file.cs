// Title: Insert the same picture into cell A1 of every worksheet in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add identical image to cell A1 on all worksheets | how to loop through worksheets and insert a picture at A1 with Aspose.Cells | batch insert PNG into each sheet's A1 cell using Aspose.Cells for .NET | C# code to add a picture anchored to A1 across multiple Excel sheets | save workbook after adding same picture to every worksheet with Aspose.Cells
// Tags: Aspose.Cells insert picture at cell coordinates | batch add same image to all worksheets C# | picture anchoring to cell using Aspose.Cells API | save modified workbook after shape insertion .NET | iterate through worksheets to add shapes Aspose

using Aspose.Cells;
using System;
using System.IO;

// The example validates the presence of an input XLSX file and a PNG image, loads the workbook with Aspose.Cells, iterates over each worksheet, and inserts the image so its top‑left corner aligns with cell A1 using the Pictures.Add method. After processing all sheets, the workbook is saved to a new file, with comprehensive exception handling for file, Aspose.Cells, and general errors.
class Program
{
    static void Main()
    {
        try
        {
            // Verify that the input workbook exists
            const string inputPath = "input.xlsx";
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input workbook not found: {inputPath}");

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Verify that the picture file exists
            const string picturePath = "image.png";
            if (!File.Exists(picturePath))
                throw new FileNotFoundException($"Picture file not found: {picturePath}");

            // Iterate through all worksheets and insert the picture anchored to cell A1
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get cell A1 in the current worksheet
                Cell cellA1 = sheet.Cells["A1"];

                // Insert the picture with its top‑left corner at cell A1
                sheet.Pictures.Add(cellA1.Row, cellA1.Column, picturePath);
            }

            // Save the modified workbook to a new file
            const string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (FileNotFoundException ex)
        {
            Console.Error.WriteLine($"File error: {ex.Message}");
        }
        catch (CellsException ex)
        {
            Console.Error.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}

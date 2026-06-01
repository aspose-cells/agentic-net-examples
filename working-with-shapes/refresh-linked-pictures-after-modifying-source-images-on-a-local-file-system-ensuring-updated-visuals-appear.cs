using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class RefreshLinkedPictures
{
    static void Main()
    {
        // Path to the workbook that contains linked pictures
        string workbookPath = "LinkedPictures.xlsx";

        try
        {
            // Verify that the source workbook exists
            if (!File.Exists(workbookPath))
                throw new FileNotFoundException($"The workbook '{workbookPath}' was not found.");

            // Load the existing workbook
            Workbook workbook = new Workbook(workbookPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each picture in the worksheet
                foreach (Picture picture in sheet.Pictures)
                {
                    // Process only linked pictures
                    if (picture.IsLink)
                    {
                        // Store the current source path
                        string sourcePath = picture.SourceFullName;

                        // Clear any cached image data (if present)
                        picture.Data = null;

                        // Reassign the source path to trigger a refresh of the linked image
                        picture.SourceFullName = sourcePath;
                    }
                }
            }

            // Save the workbook with refreshed linked pictures
            string outputPath = "LinkedPictures_Refreshed.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (FileNotFoundException fnfEx)
        {
            Console.Error.WriteLine($"File not found: {fnfEx.Message}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
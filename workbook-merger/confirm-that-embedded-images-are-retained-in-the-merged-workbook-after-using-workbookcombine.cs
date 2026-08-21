// Title: Confirm Embedded Images Survive Workbook.Combine Merge in Aspose.Cells for .NET
// Description: Demonstrates how to embed a PNG into a cell, merge the source workbook into a destination workbook with Workbook.Combine, save and reload the file, then use Cells.GetCellsWithPlaceInCellPicture to verify that the embedded image remains after the merge.
// Keywords: Aspose.Cells | Workbook.Combine | embedded image | PlaceInCell picture | C# | preserve pictures after merge | GetCellsWithPlaceInCellPicture | verify image retention | Excel workbook merge | Aspose.Cells .NET example
// Common Searches: keep embedded images when combining workbooks Aspose.Cells | Workbook.Combine image retention | how to detect embedded pictures after workbook merge | GetCellsWithPlaceInCellPicture after combine | Aspose.Cells merge workbooks preserve pictures
// Developer Intent: Validate that images embedded in cells are retained after merging workbooks with Workbook.Combine.
// Use Cases: Merging multiple report workbooks while preserving cell‑level images. | Automated quality check that embedded pictures survive workbook consolidation. | Generating a consolidated Excel file from templates that contain logos or icons placed inside cells.
// AI Prompts: Write C# code using Aspose.Cells to combine two workbooks and verify embedded images are still present. | Create a unit test that asserts the count of embedded images after Workbook.Combine matches the source workbook. | Explain how Cells.GetCellsWithPlaceInCellPicture can be used to enumerate embedded pictures after a workbook merge.

using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEmbeddedImageCombineDemo
{
    // Demonstrates how to embed a PNG into a cell, merge the source workbook into a destination workbook with Workbook.Combine, save and reload the file, then use Cells.GetCellsWithPlaceInCellPicture to verify that the embedded image remains after the merge.
    class Program
    {
        static void Main()
        {
            try
            {
                // Prepare a simple PNG image (1x1 pixel) as a byte array
                const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK0cAAAAASUVORK5CYII=";
                byte[] imageBytes = Convert.FromBase64String(base64Png);

                // ---------- Create source workbook with an embedded image ----------
                Workbook sourceWorkbook = new Workbook();
                Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

                // Embed the image into cell B2 (row 1, column 1)
                sourceSheet.Cells["B2"].EmbeddedImage = imageBytes;

                // ---------- Create destination workbook ----------
                Workbook destWorkbook = new Workbook();
                Worksheet destSheet = destWorkbook.Worksheets[0];
                destSheet.Cells["A1"].PutValue("Destination Workbook");

                // ---------- Combine source workbook into destination workbook ----------
                destWorkbook.Combine(sourceWorkbook);

                // Save the combined workbook
                string combinedPath = "CombinedWorkbook.xlsx";
                destWorkbook.Save(combinedPath, SaveFormat.Xlsx);

                // Reload the combined workbook to verify persistence of embedded images
                if (File.Exists(combinedPath))
                {
                    Workbook reloadedWorkbook = new Workbook(combinedPath);
                    Worksheet reloadedSheet = reloadedWorkbook.Worksheets[0];
                    Cells cells = reloadedSheet.Cells;

                    // Enumerate cells that contain embedded pictures (PlaceInCell)
                    int embeddedImageCount = 0;
                    IEnumerator enumerator = cells.GetCellsWithPlaceInCellPicture();
                    while (enumerator != null && enumerator.MoveNext())
                    {
                        Cell cell = (Cell)enumerator.Current;
                        if (cell.EmbeddedImage != null && cell.EmbeddedImage.Length > 0)
                        {
                            embeddedImageCount++;
                            Console.WriteLine($"Embedded image found in cell {cell.Name}");
                        }
                    }

                    Console.WriteLine($"Total embedded images after combine: {embeddedImageCount}");
                }
                else
                {
                    Console.WriteLine($"Combined file not found at path: {combinedPath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

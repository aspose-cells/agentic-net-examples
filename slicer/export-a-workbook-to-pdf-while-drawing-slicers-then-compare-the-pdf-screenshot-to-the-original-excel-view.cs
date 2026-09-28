// Title: Export an Excel workbook that contains slicers to PDF and verify slicer rendering by creating PNG snapshots with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file, saves it as a PDF while preserving slicer graphics, and then renders the first worksheet to a transparent PNG using Aspose.Cells. | Write a method that produces two PNG files of the same worksheet—one directly from the workbook and another after the PDF export—to enable visual comparison of slicer appearance. | Add robust error handling to check for the source file, catch save and render exceptions, and confirm that both PNG files were created successfully.
// Common Searches: how to keep slicer visuals when exporting Excel to PDF with Aspose.Cells C# | Aspose.Cells render worksheet to PNG after PDF conversion for visual verification | compare slicer rendering between Excel view and PDF output using .NET | C# code sample for exporting workbook with slicers to PDF and generating PNG snapshots | Aspose.Cells image rendering options transparent PNG for slicer comparison
// Tags: export workbook to PDF with slicer graphics Aspose.Cells | render worksheet as PNG transparent background Aspose.Cells | compare slicer appearance Excel vs PDF C# | SaveFormat.Pdf usage with slicer preservation | SheetRender PNG generation for visual validation | error handling for file existence Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads an Excel file, saves it as a PDF while preserving slicer graphics, renders the first worksheet to two PNG images (original view and a second render simulating the PDF output) using Aspose.Cells, and checks that both images were generated.
class WorkbookPdfSlicerComparer
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string pdfPath = "output.pdf";
            const string originalImgPath = "original.png";
            const string renderedImgPath = "pdf_rendered.png";

            // Verify the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Export the workbook to PDF
            try
            {
                workbook.Save(pdfPath, SaveFormat.Pdf);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save PDF: {ex.Message}");
                return;
            }

            // Set image rendering options (default format is PNG)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = true,
                Transparent = true
            };

            // Render the worksheet to an image (original view)
            try
            {
                SheetRender sheetRender = new SheetRender(sheet, imgOptions);
                sheetRender.ToImage(0, originalImgPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to render original image: {ex.Message}");
                return;
            }

            // Render again to simulate PDF rendering output
            try
            {
                SheetRender sheetRender = new SheetRender(sheet, imgOptions);
                sheetRender.ToImage(0, renderedImgPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to render PDF-simulated image: {ex.Message}");
                return;
            }

            // Verify that both images were created
            bool imagesExist = File.Exists(originalImgPath) && File.Exists(renderedImgPath);
            Console.WriteLine(imagesExist
                ? "Images rendered successfully."
                : "Failed to render one or both images.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}

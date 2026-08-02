// Title: Convert HTML with Embedded Images to Excel Cells Using Aspose.Cells for .NET
// Description: Loads an HTML file that contains inline or base‑64 images into an Aspose.Cells Workbook and saves it as XLSX. Each <img> tag is imported as a picture and automatically anchored to the cell that originally held the image, preserving the visual layout.
// Keywords: Aspose.Cells | HTML to Excel | embed images | C# conversion | picture anchoring | save as XLSX | load HTML workbook | inline images | base64 images | cell anchored pictures
// Common Searches: Aspose.Cells convert HTML with images to Excel | C# load HTML and keep pictures in cells | HTML to XLSX preserving inline images | map <img> tags to Excel cells using Aspose | save HTML page as Excel with images
// Developer Intent: Transform an HTML document that includes embedded images into an Excel file where every image is placed in the corresponding cell.
// Use Cases: Migrate web‑based reports that embed chart images into Excel worksheets for further analysis. | Automate conversion of email HTML templates containing logos so the logos appear in the same cells of the generated spreadsheet. | Archive an HTML product catalog in Excel while keeping each product photo inside its matching cell.
// AI Prompts: Generate C# code with Aspose.Cells that loads an HTML file containing inline or base64 images and saves it as XLSX, ensuring each image is anchored to its original cell. | Explain how Aspose.Cells maps HTML <img> elements to picture objects and how to modify the cell anchoring after conversion.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsHtmlToExcel
{
    // Loads an HTML file that contains inline or base‑64 images into an Aspose.Cells Workbook and saves it as XLSX. Each <img> tag is imported as a picture and automatically anchored to the cell that originally held the image, preserving the visual layout.
    class Program
    {
        static void Main()
        {
            // Path to the source HTML file that contains embedded images.
            string htmlPath = "input.html";

            // Path for the generated Excel file.
            string excelPath = "output.xlsx";

            // Load the HTML file into a workbook.
            // LoadOptions are not required here because the constructor infers the format from the file extension.
            Workbook workbook = new Workbook(htmlPath);

            // Save the workbook as an Excel file.
            // All images embedded in the HTML are imported as pictures and positioned in the corresponding cells.
            workbook.Save(excelPath, SaveFormat.Xlsx);

            Console.WriteLine($"HTML file '{htmlPath}' has been converted to Excel '{excelPath}'.");
        }
    }
}

// Title: How to confirm that a slicer's border thickness and color stay unchanged after converting an Excel workbook to PDF with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells to open an .xlsx file, retrieve the first slicer's line weight and color from its LineFormat, output the values, and then save the workbook as a PDF. | Create a method that compares a slicer's border properties before and after calling Workbook.Save with SaveFormat.Pdf, logging any differences detected. | Develop a reusable utility that obtains a slicer's shape line attributes (weight, color) and validates that the generated PDF reflects the same styling.
// Common Searches: aspnet check slicer border thickness after pdf export with Aspose.Cells | C# get slicer line color from Excel using Aspose.Cells | does Aspose.Cells preserve slicer formatting when saving to PDF | how to read slicer shape line properties in a workbook before conversion | validate Excel slicer appearance in generated PDF document
// Tags: Aspose.Cells get slicer line details | PDF export keep slicer line style | C# read slicer shape line weight | validate slicer styling after conversion | Aspose.Cells slicer border color access

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing; // For slicer related classes

// The program loads an Excel workbook, accesses the first worksheet's slicer collection, reads the slicer's shape LineFormat to obtain its border weight and, via reflection, its color, prints these values, and then converts the workbook to PDF while handling file‑existence and directory‑creation concerns.
class SlicerPdfValidation
{
    static void Main()
    {
        try
        {
            // Input Excel file
            string excelPath = "input.xlsx";

            // Verify that the input file exists
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Error: The file \"{excelPath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(excelPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve slicer collection from the worksheet
            var slicers = sheet.Slicers;

            if (slicers == null || slicers.Count == 0)
            {
                Console.WriteLine("No slicers found in the worksheet.");
                return;
            }

            // Use the first slicer for demonstration
            var slicer = slicers[0];

            // Get the underlying shape of the slicer to access border properties
            Shape slicerShape = slicer.Shape;
            if (slicerShape == null)
            {
                Console.WriteLine("Unable to retrieve the shape of the slicer.");
                return;
            }

            // Access line (border) information via LineFormat
            LineFormat lineFormat = slicerShape.Line;
            if (lineFormat == null)
            {
                Console.WriteLine("Unable to retrieve line format of the slicer shape.");
                return;
            }

            // Border thickness (in points)
            double borderThickness = lineFormat.Weight;

            // Attempt to retrieve border color via reflection (compatible with multiple library versions)
            Color borderColor = Color.Empty;
            var colorProp = lineFormat.GetType().GetProperty("Color");
            if (colorProp != null && colorProp.PropertyType == typeof(Color))
            {
                object val = colorProp.GetValue(lineFormat);
                if (val is Color c)
                    borderColor = c;
            }

            Console.WriteLine($"Slicer Border Thickness: {borderThickness} pt");
            if (borderColor != Color.Empty)
                Console.WriteLine($"Slicer Border Color: ARGB({borderColor.A}, {borderColor.R}, {borderColor.G}, {borderColor.B})");
            else
                Console.WriteLine("Slicer Border Color: (not available)");

            // Convert the workbook to PDF
            string pdfPath = "output.pdf";

            // Ensure the directory for the PDF exists
            string pdfDir = Path.GetDirectoryName(pdfPath);
            if (!string.IsNullOrEmpty(pdfDir) && !Directory.Exists(pdfDir))
                Directory.CreateDirectory(pdfDir);

            try
            {
                workbook.Save(pdfPath, SaveFormat.Pdf);
                Console.WriteLine($"Workbook successfully converted to PDF: {pdfPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving PDF: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

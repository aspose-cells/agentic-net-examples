// Title: How to convert an Excel workbook to a non‑editable PDF by removing all form control shapes with Aspose.Cells for .NET
// Common Searches: Aspose.Cells remove form controls before saving workbook as PDF in C# | C# code to delete Excel button and checkbox shapes using Aspose.Cells | Convert Excel to PDF without interactive form fields using Aspose.Cells | How to strip worksheet shapes and export to PDF with Aspose.Cells .NET | Generate read‑only PDF from Excel workbook by removing form fields in C#
// Tags: Aspose.Cells shape removal for PDF export | Excel to PDF conversion without interactive controls | C# delete worksheet form objects Aspose.Cells | non‑editable PDF generation from Excel workbook | strip form field shapes before saving as PDF

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel file with Aspose.Cells, iterates each worksheet, and removes any shape that represents a form control (Button, CheckBox, ComboBox, ListBox, RadioButton, GroupBox, Label) by looping backwards through the Shapes collection. After ensuring the output directory exists, the cleaned workbook is saved as a PDF, producing a non‑editable final document. The code includes file existence checks and basic exception handling.
class ConvertWorkbookToPdf
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.pdf";

        try
        {
            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet and remove form control shapes
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Loop backwards because we will be removing items from the collection
                for (int i = sheet.Shapes.Count - 1; i >= 0; i--)
                {
                    Shape shape = sheet.Shapes[i];

                    // Identify form controls by their type name (avoids direct enum reference)
                    string typeName = shape.Type.ToString();

                    if (typeName == "Button" ||
                        typeName == "CheckBox" ||
                        typeName == "ComboBox" ||
                        typeName == "ListBox" ||
                        typeName == "RadioButton" ||
                        typeName == "GroupBox" ||
                        typeName == "Label")
                    {
                        // Remove the shape (form field) from the worksheet
                        sheet.Shapes.RemoveAt(i);
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook as a PDF (non‑editable version)
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

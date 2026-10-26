// Title: How to embed a PDF as an OLE object in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: aspocells c# add ole object pdf to worksheet | embed a PDF file as an OLE object in Excel using Aspose.Cells | C# code to insert a PDF into an XLSX workbook as an OLE object | how to programmatically embed a PDF in an Excel sheet with Aspose.Cells for .NET | save workbook after adding PDF OLE object using Aspose.Cells
// Tags: aspocells add oleobject pdf | excel worksheet embed pdf oleobject c# | aspose.cells insert ole object into xlsx | programmatic pdf embedding in excel workbook | worksheet shapes add oleobject aspocells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example checks for the presence of an input Excel file and a PDF file, loads the workbook, reads the PDF bytes, inserts the PDF as an OLE object at a specified cell range using Aspose.Cells, and saves the updated workbook to a new file.
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

            // Verify that the PDF file to embed exists
            const string pdfPath = "document.pdf";
            if (!File.Exists(pdfPath))
                throw new FileNotFoundException($"PDF file not found: {pdfPath}");

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Read PDF bytes
            byte[] pdfData = File.ReadAllBytes(pdfPath);

            // Add the OLE object (PDF) to the worksheet.
            // Parameters: upper left row, upper left column, height (pixels), width (pixels),
            // oleObjectType (0 = unknown), oleObjectProgId (0 = unknown), OLE object data bytes
            OleObject ole = sheet.Shapes.AddOleObject(2, 2, 100, 200, 0, 0, pdfData);

            // Save the modified workbook
            const string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

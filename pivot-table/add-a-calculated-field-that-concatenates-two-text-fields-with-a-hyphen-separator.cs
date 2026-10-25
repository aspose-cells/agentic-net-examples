// Title: How to add a calculated column that concatenates two text columns with a hyphen in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add calculated column that joins two columns with a hyphen | Programmatically insert =A2&"-"&B2 formula into a worksheet with Aspose.Cells | Find next empty column and set header in an Aspose.Cells workbook | Recalculate formulas after adding a concatenated field using Aspose.Cells .NET | Create or load an Excel file and add a hyphen‑separated concatenated column with Aspose.Cells
// Tags: calculated column insertion Aspose.Cells | text concatenation formula Excel C# | next empty column detection Aspose.Cells | workbook formula recalculation Aspose.Cells | save workbook output.xlsx Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing workbook or creates a new one, identifies the next empty column, adds a header named "Concatenated", and writes a formula that joins the values from columns A and B with a hyphen (e.g., =A2&"-"&B2) for each data row. After inserting the formulas it recalculates the workbook and saves the result to output.xlsx.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one with sample data
            if (File.Exists(inputPath))
            {
                // LoadOptions can be used to specify a password if needed; leaving it empty for unprotected files
                var loadOptions = new LoadOptions();
                workbook = new Workbook(inputPath, loadOptions);
            }
            else
            {
                workbook = new Workbook();
                Worksheet sampleSheet = workbook.Worksheets[0];
                // Sample headers
                sampleSheet.Cells["A1"].PutValue("First");
                sampleSheet.Cells["B1"].PutValue("Second");
                // Sample data rows
                sampleSheet.Cells["A2"].PutValue("Hello");
                sampleSheet.Cells["B2"].PutValue("World");
                sampleSheet.Cells["A3"].PutValue("Foo");
                sampleSheet.Cells["B3"].PutValue("Bar");
            }

            Worksheet sheet = workbook.Worksheets[0];

            // Determine the next empty column index
            int newColIndex = sheet.Cells.MaxColumn + 1;

            // Set header for the new calculated field
            sheet.Cells[0, newColIndex].PutValue("Concatenated");

            // Get the last row that contains data (excluding header)
            int lastDataRow = sheet.Cells.MaxDataRow;

            // Apply concatenation formula to each data row
            for (int row = 1; row <= lastDataRow; row++)
            {
                // Formula: =A{row+1}&"-"&B{row+1}
                string formula = $"=A{row + 1}&\"-\"&B{row + 1}";
                sheet.Cells[row, newColIndex].Formula = formula;
            }

            // Recalculate formulas so that values are stored
            workbook.CalculateFormula();

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}

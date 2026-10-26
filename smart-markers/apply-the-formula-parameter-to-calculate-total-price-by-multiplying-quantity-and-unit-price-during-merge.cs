// Title: Calculate total price with a multiplication formula in a merged title row using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set formula for total price column after merging header row | multiply quantity by unit price in Excel using Aspose.Cells .NET | apply formula to merged cells in an Aspose.Cells workbook C# example | how to recalculate formulas after adding a multiplication formula with Aspose.Cells | create sales report with merged title and calculated total price using Aspose.Cells
// Tags: multiplication formula in Aspose.Cells C# | merged title row Excel Aspose.Cells | derived column calculation Aspose.Cells | formula recalculation workbook Aspose.Cells | export workbook to xlsx Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Creates a new workbook, adds headers and sample data, merges the first row for a title, assigns a multiplication formula (Quantity * Unit Price) to the Total Price column, recalculates formulas, and saves the file as SalesReport.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add headers
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["C1"].PutValue("Unit Price");
            sheet.Cells["D1"].PutValue("Total Price");

            // Sample data
            sheet.Cells["A2"].PutValue("Widget");
            sheet.Cells["B2"].PutValue(5);          // Quantity
            sheet.Cells["C2"].PutValue(12.99);      // Unit Price

            sheet.Cells["A3"].PutValue("Gadget");
            sheet.Cells["B3"].PutValue(3);
            sheet.Cells["C3"].PutValue(7.5);

            // Merge cells for a title (optional demonstration)
            sheet.Cells.Merge(0, 0, 1, 4); // Merge A1:D1 for a title row
            sheet.Cells["A1"].PutValue("Sales Report");

            // Center align the title using Style object
            Style titleStyle = workbook.CreateStyle();
            titleStyle.HorizontalAlignment = TextAlignmentType.Center;
            sheet.Cells["A1"].SetStyle(titleStyle);

            // Apply formulas to calculate Total Price = Quantity * Unit Price
            sheet.Cells["D2"].Formula = "B2*C2";
            sheet.Cells["D3"].Formula = "B3*C3";

            // Recalculate the workbook to evaluate formulas
            workbook.CalculateFormula();

            // Define output file path
            string outputPath = "SalesReport.xlsx";

            // Ensure the directory exists (in case a relative path is used)
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

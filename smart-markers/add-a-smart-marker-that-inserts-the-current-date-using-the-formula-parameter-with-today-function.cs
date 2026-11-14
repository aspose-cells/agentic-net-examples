// Title: How to insert the current date into an Excel worksheet using Aspose.Cells smart marker with the TODAY() formula in C#
// Common Searches: C# Aspose.Cells smart marker to display today's date in Excel | Using WorkbookDesigner to evaluate TODAY() formula in a smart marker | Insert dynamic date with smart marker formula in Aspose.Cells C# example | Aspose.Cells smart marker syntax for Formula(TODAY()) in C# | Generate Excel file with current date using Aspose.Cells smart markers
// Tags: smart-marker formula today c# | Aspose.Cells dynamic date insertion | WorkbookDesigner evaluate smart marker formulas | generate excel file with current date | excel smart marker date placeholder

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, places a smart marker «CurrentDate:Formula(TODAY())» in cell A1, processes it with WorkbookDesigner to evaluate the TODAY() function, and saves the result as SmartMarkerDate.xlsx.
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

            // Insert a smart marker that uses the Formula parameter with TODAY() to insert the current date
            // Smart marker syntax: «FieldName:Formula(Expression)»
            sheet.Cells["A1"].PutValue("«CurrentDate:Formula(TODAY())»");

            // Process the smart markers using WorkbookDesigner (no data source required for formula evaluation)
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.Process();

            // Define output file path
            string outputPath = "SmartMarkerDate.xlsx";

            // Save the resulting workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

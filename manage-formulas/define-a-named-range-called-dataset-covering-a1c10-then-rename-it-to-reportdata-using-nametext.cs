using System;
using Aspose.Cells;

class DefineAndRenameNamedRange
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Create a range that covers cells A1:C10 (use fully qualified type to avoid ambiguity)
            Aspose.Cells.Range range = sheet.Cells.CreateRange("A1:C10");

            // Assign the initial name "DataSet" to the range
            range.Name = "DataSet";

            // Retrieve the Name object that was created by the previous step
            Name namedRange = workbook.Worksheets.Names["DataSet"];

            // Rename the range to "ReportData" using the Text property
            namedRange.Text = "ReportData";

            // Save the workbook to a file
            string outputPath = "NamedRangeRenamed.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
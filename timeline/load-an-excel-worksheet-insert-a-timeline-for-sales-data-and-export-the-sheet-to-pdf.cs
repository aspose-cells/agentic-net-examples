using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Saving;

class TimelineToPdfDemo
{
    static void Main()
    {
        try
        {
            const string inputFile = "SalesData.xlsx";
            const string outputFile = "SalesReport.pdf";

            // Verify that the source workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file \"{inputFile}\" not found.");
                return;
            }

            // Load the workbook that contains the sales data
            Workbook workbook = new Workbook(inputFile);
            Worksheet sheet = workbook.Worksheets[0];

            // Define the source range for the pivot table (adjust as needed)
            // Assuming the first row contains headers "Date" and "Sales"
            string sourceRange = "A1:B10";

            // Add a pivot table starting at cell D1
            int pivotIndex = sheet.PivotTables.Add(sourceRange, "D1", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Refresh and calculate the pivot table
            pivot.RefreshData();
            pivot.CalculateData();

            // Ensure the pivot contains the required row field before adding a Timeline
            bool hasDateField = false;
            foreach (PivotField field in pivot.RowFields)
            {
                if (field.Name.Equals("Date", StringComparison.OrdinalIgnoreCase))
                {
                    hasDateField = true;
                    break;
                }
            }

            if (hasDateField)
            {
                // Insert a Timeline linked to the pivot table at cell F1
                sheet.Timelines.Add(pivot, "F1", "Date");
            }
            else
            {
                Console.WriteLine("The pivot table does not contain a \"Date\" row field; Timeline will not be added.");
            }

            // Save the workbook (including the Timeline) as PDF
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            workbook.Save(outputFile, pdfOptions);

            Console.WriteLine($"PDF report generated successfully: {outputFile}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
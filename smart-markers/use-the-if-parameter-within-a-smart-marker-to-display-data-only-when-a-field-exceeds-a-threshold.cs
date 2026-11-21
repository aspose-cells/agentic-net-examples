// Title: Use IF smart marker to display a field only when its value exceeds a threshold in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# smart marker IF condition to display values greater than 80 | How to hide Excel cell content unless numeric value exceeds a threshold using Aspose.Cells | C# example of conditional smart marker processing with WorkbookDesigner | Smart marker syntax {IF:Score>80} in Aspose.Cells .NET | Generate Excel report with conditional rows based on score using Aspose.Cells
// Tags: Aspose.Cells IF smart marker syntax | WorkbookDesigner conditional processing C# | Excel threshold filtering with smart markers | DataTable as smart marker data source | Conditional content generation Aspose.Cells

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The example creates an in‑memory workbook, inserts an IF smart marker that shows the Score only when it is greater than 80, fills a DataTable with score values, processes the markers with WorkbookDesigner, ensures the output directory exists, and saves the result as SmartMarkerIfResult.xlsx.
class SmartMarkerIfExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook (template) in memory
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a smart marker with IF condition.
            // Displays the score only when it exceeds the threshold (80).
            // Syntax: {IF:Score>80}Score: {Score}{ENDIF}
            sheet.Cells["A1"].PutValue("{IF:Score>80}Score: {Score}{ENDIF}");

            // Prepare data source
            DataTable dt = new DataTable("Data");
            dt.Columns.Add("Score", typeof(int));
            dt.Rows.Add(75); // Below threshold, will not be displayed
            dt.Rows.Add(85); // Above threshold, will be displayed
            dt.Rows.Add(90); // Above threshold, will be displayed

            // Process smart markers using WorkbookDesigner (compatible with all Aspose.Cells versions)
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dt);
            designer.Process();

            // Define output file path
            string outputPath = "SmartMarkerIfResult.xlsx";

            // Ensure the directory exists before saving (handle cases where outputPath has no directory part)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the result
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}

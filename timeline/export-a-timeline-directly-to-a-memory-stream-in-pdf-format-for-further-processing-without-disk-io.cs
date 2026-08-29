// Title: Export an Aspose.Cells timeline to a PDF using a MemoryStream in C#
// AI Prompts: Write C# code that builds a pivot table, attaches a timeline, and saves the workbook as a PDF into a MemoryStream with Aspose.Cells. | Show how to reset the MemoryStream position after saving the PDF so it can be consumed by other components.
// Common Searches: Aspose.Cells C# export timeline to PDF without writing a file | How to generate an in‑memory PDF from an Excel timeline using Aspose.Cells | Saving a workbook with a timeline to a MemoryStream as PDF in .NET | Read PDF bytes from MemoryStream after Aspose.Cells SaveFormat.Pdf | C# Aspose.Cells timeline PDF output stream example
// Tags: Aspose.Cells timeline PDF stream export | C# in‑memory PDF generation Aspose.Cells | pivot table timeline PDF save Aspose.Cells | save workbook as PDF stream .NET | memory stream PDF output Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

// The example creates a workbook, adds sample data, builds a pivot table, links a timeline, and then saves the entire workbook—including the timeline—as a PDF directly into a MemoryStream, resets the stream position, and prints the generated PDF size.
public class ExportTimelineToPdfMemoryStream
{
    public static void Run()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate worksheet with sample data including a date field
            worksheet.Cells["A1"].PutValue("Date");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["A2"].PutValue(DateTime.Now.AddDays(-4));
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["A3"].PutValue(DateTime.Now.AddDays(-3));
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["A4"].PutValue(DateTime.Now.AddDays(-2));
            worksheet.Cells["B4"].PutValue(30);
            worksheet.Cells["A5"].PutValue(DateTime.Now.AddDays(-1));
            worksheet.Cells["B5"].PutValue(40);
            worksheet.Cells["A6"].PutValue(DateTime.Now);
            worksheet.Cells["B6"].PutValue(50);

            // Create a pivot table that will serve as the data source for the timeline
            int pivotIndex = worksheet.PivotTables.Add("A1:B6", "D1", "PivotTable1");
            PivotTable pivot = worksheet.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");

            // Refresh pivot cache data using the correct API
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a timeline linked to the pivot table
            int timelineIndex = worksheet.Timelines.Add(pivot, "F1", "Date");
            Timeline timeline = worksheet.Timelines[timelineIndex];
            timeline.Caption = "Sample Timeline";

            // Create a memory stream to hold the PDF output
            using (MemoryStream pdfStream = new MemoryStream())
            {
                // Save the workbook (including the timeline) as PDF into the memory stream
                workbook.Save(pdfStream, SaveFormat.Pdf);

                // Reset the stream position for any subsequent reading
                pdfStream.Position = 0;

                // Output the size of the generated PDF
                Console.WriteLine($"Generated PDF stream length: {pdfStream.Length} bytes");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Entry point for the application
    public static void Main()
    {
        Run();
    }
}

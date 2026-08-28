// Title: Apply French (fr-FR) CultureInfo to Aspose.Cells smart markers for date and number localization in C#
// AI Prompts: Write C# code that assigns Workbook.Settings.CultureInfo to a specific locale before processing smart markers with WorkbookDesigner. | Show how to bind a DataTable to WorkbookDesigner and let the workbook's CultureInfo control the formatting of date and numeric smart markers. | Explain the effect of changing the workbook's CultureInfo on smart marker output and demonstrate saving the localized Excel file.
// Common Searches: Aspose.Cells C# set workbook culture to French for smart marker formatting | How to localize dates in smart markers using CultureInfo in Aspose.Cells | Smart marker number formatting based on workbook locale in .NET | WorkbookDesigner data source binding with French locale example
// Tags: Workbook.Settings.CultureInfo localization Aspose.Cells | smart marker date formatting C# | smart marker number formatting with locale | WorkbookDesigner data source binding example | export Excel with French culture Aspose

using System;
using System.Data;
using System.Globalization;
using Aspose.Cells;

// The example creates a workbook, inserts smart markers for a date and a number, populates a DataTable, sets the workbook's CultureInfo to French (fr-FR) to control formatting, binds the data source to WorkbookDesigner, processes the smart markers, and saves the localized Excel file.
public class SmartMarkerLocaleDemo
{
    public static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Insert smart markers that will be replaced with data values
            ws.Cells["A1"].PutValue("Date: <#=Date#>");
            ws.Cells["A2"].PutValue("Number: <#=Number#>");

            // Prepare a data source with a date and a numeric value
            DataTable dt = new DataTable("Data");
            dt.Columns.Add("Date", typeof(DateTime));
            dt.Columns.Add("Number", typeof(double));
            dt.Rows.Add(DateTime.Now, 12345.67);

            // Set workbook culture to French (fr-FR) to affect formatting
            wb.Settings.CultureInfo = new CultureInfo("fr-FR");

            // Use WorkbookDesigner to process smart markers
            WorkbookDesigner designer = new WorkbookDesigner(wb);
            designer.SetDataSource(dt);

            // Process the smart markers with the provided data and locale
            designer.Process();

            // Save the resulting workbook
            string outputPath = "SmartMarkerLocaleDemo.xlsx";
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

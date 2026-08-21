// Title: Aspose.Cells for .NET – Localize PivotTable Row Header Caption (RowHeaderCaption)
// Description: Demonstrates how to create a workbook, add sample data, build a PivotTable on a separate sheet, and set the RowHeaderCaption to a localized string (e.g., Chinese "行标题"). The example configures row and data fields and saves the result as an XLSX file.
// Keywords: Aspose.Cells | .NET | C# | PivotTable | RowHeaderCaption | localize pivot header | internationalization | i18n | multilingual Excel | Chinese Excel header | Japanese PivotTable label | Spanish workbook localization | Excel automation | XLSX export | sample code
// Common Searches: Aspose.Cells set PivotTable RowHeaderCaption | localize PivotTable row header in C# | change PivotTable row header language Aspose.Cells | RowHeaderCaption Chinese example | internationalize Excel PivotTable labels with Aspose | how to use RowHeaderCaption property .NET | multilingual PivotTable Aspose.Cells tutorial
// Developer Intent: Apply a localized string to the PivotTable.RowHeaderCaption property in an Aspose.Cells workbook.
// Use Cases: Generate financial reports with row headers displayed in the end‑user's language (e.g., Chinese, Japanese, Spanish). | Create Excel workbooks for global markets where PivotTable labels must match the application's culture settings. | Automate workbook production pipelines that adapt PivotTable captions based on .resx resources or runtime culture.
// AI Prompts: Show C# code that reads a row header caption from a .resx file and assigns it to PivotTable.RowHeaderCaption using Aspose.Cells. | Generate a function that sets PivotTable.RowHeaderCaption dynamically based on Thread.CurrentThread.CurrentUICulture. | Explain step‑by‑step how to internationalize all PivotTable captions (RowHeaderCaption, ColumnHeaderCaption, etc.) in an Aspose.Cells workbook. | Provide a GitHub‑style README snippet describing how to localize PivotTable headers for multiple languages with Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Pivot;

// Demonstrates how to create a workbook, add sample data, build a PivotTable on a separate sheet, and set the RowHeaderCaption to a localized string (e.g., Chinese "行标题"). The example configures row and data fields and saves the result as an XLSX file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the default worksheet for data
        Workbook workbook = new Workbook();
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Name = "Data";

        // Populate sample data for the pivot table
        dataSheet.Cells["A1"].PutValue("Category");
        dataSheet.Cells["B1"].PutValue("Amount");
        dataSheet.Cells["A2"].PutValue("Food");
        dataSheet.Cells["B2"].PutValue(120);
        dataSheet.Cells["A3"].PutValue("Drink");
        dataSheet.Cells["B3"].PutValue(80);

        // Add a separate worksheet that will contain the pivot table
        Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

        // Create the pivot table (source range, destination cell, pivot name)
        int pivotIndex = pivotSheet.PivotTables.Add("=Data!A1:B3", "A3", "MyPivot");
        PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

        // Set the Row Header caption to a localized string (e.g., Chinese)
        pivotTable.RowHeaderCaption = "行标题";

        // Configure the pivot fields
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

        // Save the workbook to a file
        workbook.Save("PivotRowHeaderLocalized.xlsx", SaveFormat.Xlsx);
    }
}

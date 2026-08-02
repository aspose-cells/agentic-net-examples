// Title: Refresh a linked SmartArt shape after changing its source cell with Aspose.Cells for .NET
// Description: Shows how to open an Excel workbook, update a cell that drives a SmartArt diagram, call Worksheet.Shapes.UpdateSelectedValue (or Shape.RefreshLinkedShape) to refresh the linked SmartArt, enable automatic diagram updates via OoxmlSaveOptions.UpdateSmartArt, and save the file so the visual reflects the new data.
// Keywords: Aspose.Cells | .NET | SmartArt refresh | linked SmartArt | UpdateSmartArt | Worksheet.Shapes.UpdateSelectedValue | Shape.RefreshLinkedShape | Excel automation | programmatic SmartArt update | OoxmlSaveOptions
// Common Searches: Aspose.Cells refresh linked SmartArt | Update SmartArt after cell change C# | Shape.RefreshLinkedShape example | Enable SmartArt update when saving workbook | How to programmatically refresh SmartArt in Excel
// Developer Intent: Refresh a SmartArt diagram that is linked to a worksheet cell after the cell value is modified in code.
// Use Cases: Automated KPI dashboards where SmartArt graphics must reflect the latest cell values before distribution. | Scheduled report generation that modifies data cells and needs the associated SmartArt to update automatically. | Processing Excel templates with linked SmartArt, altering source cells via code, and delivering a final file with synchronized diagrams.
// AI Prompts: Provide a C# example that changes a cell value, refreshes the linked SmartArt, and saves the workbook with UpdateSmartArt enabled. | Explain when to use Worksheet.Shapes.UpdateSelectedValue versus Shape.RefreshLinkedShape in Aspose.Cells. | How can I ensure SmartArt diagrams update automatically when I modify their source cells using Aspose.Cells for .NET?

using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to open an Excel workbook, update a cell that drives a SmartArt diagram, call Worksheet.Shapes.UpdateSelectedValue (or Shape.RefreshLinkedShape) to refresh the linked SmartArt, enable automatic diagram updates via OoxmlSaveOptions.UpdateSmartArt, and save the file so the visual reflects the new data.
class RefreshSmartArtDemo
{
    static void Main()
    {
        // Load the workbook that contains the linked SmartArt shape
        Workbook workbook = new Workbook("template.xlsx");

        // Update the cell that is linked to the SmartArt shape
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["B2"].PutValue("New Value");

        // Refresh the selected values of all shapes (including linked SmartArt)
        worksheet.Shapes.UpdateSelectedValue();

        // Save the workbook with SmartArt update enabled
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.UpdateSmartArt = true;
        workbook.Save("output.xlsx", saveOptions);
    }
}

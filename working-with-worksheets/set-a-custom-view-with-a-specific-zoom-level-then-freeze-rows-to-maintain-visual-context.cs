using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Set the view type (optional, ensures normal view)
        worksheet.ViewType = ViewType.NormalView;

        // Set a custom zoom level (e.g., 150%)
        worksheet.Zoom = 150;

        // Freeze the first 4 rows while keeping columns unfrozen
        // Parameters: row index, column index, number of frozen rows, number of frozen columns
        worksheet.FreezePanes(4, 0, 4, 0);

        // Save the workbook
        workbook.Save("CustomViewAndFreeze.xlsx");
    }
}
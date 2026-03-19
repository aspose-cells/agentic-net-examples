using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the workbook settings
        WorkbookSettings settings = workbook.Settings;

        // Hide both horizontal and vertical scrollbars for a cleaner UI
        settings.IsHScrollBarVisible = false;
        settings.IsVScrollBarVisible = false;

        // Save the workbook
        workbook.Save("CleanUIWorkbook.xlsx", SaveFormat.Xlsx);
    }
}
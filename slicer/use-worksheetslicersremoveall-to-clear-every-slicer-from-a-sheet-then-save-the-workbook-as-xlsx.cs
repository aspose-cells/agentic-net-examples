using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet (or any specific worksheet)
        Worksheet worksheet = workbook.Worksheets[0];

        // Clear all slicers from this worksheet
        worksheet.Slicers.Clear();

        // Save the workbook as an XLSX file
        workbook.Save("ClearedSlicers.xlsx", SaveFormat.Xlsx);
    }
}
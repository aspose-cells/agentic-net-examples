using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a hyperlink to cell H2 (row 2, column H) pointing to the external website
        int index = worksheet.Hyperlinks.Add("H2", 1, 1, "https://example.com");

        // Set the text that will be displayed in the cell
        worksheet.Hyperlinks[index].TextToDisplay = "Example Site";

        // Save the workbook
        workbook.Save("Hyperlink_H2.xlsx");
    }
}
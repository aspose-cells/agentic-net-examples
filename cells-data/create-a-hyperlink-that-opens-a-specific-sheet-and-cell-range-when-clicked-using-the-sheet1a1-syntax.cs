using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet (default name is Sheet1)
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Sheet1";

        // Add a second worksheet that will be the hyperlink target
        Worksheet targetSheet = workbook.Worksheets.Add("TargetSheet");
        targetSheet.Cells["A1"].PutValue("Destination cell");

        // Add a hyperlink in Sheet1 cell B2 that opens TargetSheet!A1
        // The address uses the '#SheetName!Cell' syntax for internal links
        int hyperlinkIndex = sheet.Hyperlinks.Add("B2", 1, 1, "#TargetSheet!A1");

        // Set the text that will be displayed in the cell
        sheet.Hyperlinks[hyperlinkIndex].TextToDisplay = "Go to TargetSheet A1";

        // Save the workbook
        workbook.Save("HyperlinkInternal.xlsx");
    }
}
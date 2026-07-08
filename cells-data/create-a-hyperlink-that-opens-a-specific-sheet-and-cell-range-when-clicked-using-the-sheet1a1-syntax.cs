using Aspose.Cells;

class HyperlinkToSheetCell
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a second worksheet that will be the hyperlink target
        Worksheet targetSheet = workbook.Worksheets[workbook.Worksheets.Add()];
        targetSheet.Name = "Sheet2";
        targetSheet.Cells["A1"].PutValue("Target Cell in Sheet2");

        // Get the first worksheet where the hyperlink will be placed
        Worksheet sourceSheet = workbook.Worksheets[0];
        sourceSheet.Name = "Sheet1";

        // Add a hyperlink in cell B2 that opens Sheet2!A1 using the '#Sheet2!A1' syntax
        int hyperlinkIndex = sourceSheet.Hyperlinks.Add("B2", 1, 1, "#Sheet2!A1");

        // Set the text that will be displayed for the hyperlink
        sourceSheet.Hyperlinks[hyperlinkIndex].TextToDisplay = "Go to Sheet2 A1";

        // Save the workbook
        workbook.Save("HyperlinkToSheetCell.xlsx");
    }
}
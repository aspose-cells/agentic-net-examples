using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet and give it a name
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "Sheet1";

        // Add a second worksheet and name it
        int sheet2Index = workbook.Worksheets.Add();
        Worksheet sheet2 = workbook.Worksheets[sheet2Index];
        sheet2.Name = "Sheet2";

        // Put some content in the target cell of Sheet2
        sheet2.Cells["A1"].PutValue("Destination");

        // Add a hyperlink in Sheet1!A1 that points to Sheet2!A1
        // Using HyperlinkCollection.Add(string cellName, int totalRows, int totalColumns, string address)
        sheet1.Hyperlinks.Add("A1", 1, 1, "Sheet2!A1");

        // Set the display text for the hyperlink (optional)
        sheet1.Hyperlinks[0].TextToDisplay = "Go to Sheet2";

        // Save the workbook
        workbook.Save("InternalHyperlinkDemo.xlsx");
    }
}
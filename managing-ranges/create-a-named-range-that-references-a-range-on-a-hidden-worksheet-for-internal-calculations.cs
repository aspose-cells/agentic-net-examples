using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a hidden worksheet that will hold internal data
        Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenData");
        hiddenSheet.IsVisible = false; // Hide the worksheet

        // Populate the hidden worksheet with some values
        hiddenSheet.Cells["A1"].PutValue(10);
        hiddenSheet.Cells["A2"].PutValue(20);
        hiddenSheet.Cells["A3"].PutValue(30);

        // Add a visible worksheet for user interaction (optional)
        Worksheet mainSheet = workbook.Worksheets[0];
        mainSheet.Name = "Main";

        // Create a named range that refers to the range on the hidden worksheet
        int nameIndex = workbook.Worksheets.Names.Add("CalcRange");
        Name calcRange = workbook.Worksheets.Names[nameIndex];
        calcRange.RefersTo = "=HiddenData!$A$1:$A$3"; // Reference must start with '='
        calcRange.IsVisible = false; // Keep the name hidden as well (optional)

        // Use the named range in a formula on the visible sheet
        mainSheet.Cells["B1"].Formula = "=SUM(CalcRange)";
        workbook.CalculateFormula();

        // Save the workbook
        workbook.Save("HiddenNamedRange.xlsx");
    }
}
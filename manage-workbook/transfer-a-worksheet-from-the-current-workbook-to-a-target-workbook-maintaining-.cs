using Aspose.Cells;

class TransferWorksheet
{
    static void Main()
    {
        // Create a source workbook and populate it with sample data and a formula
        Workbook sourceWorkbook = new Workbook();
        Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
        sourceSheet.Name = "SourceSheet";
        sourceSheet.Cells["A1"].PutValue("Sample");
        sourceSheet.Cells["B1"].Formula = "=A1&\"_Copy\""; // formula to be preserved

        // Create a destination workbook (initially contains one default worksheet)
        Workbook destWorkbook = new Workbook();
        Worksheet destSheet = destWorkbook.Worksheets[0];
        destSheet.Name = "DestinationSheet";

        // Copy the source worksheet into the destination worksheet.
        // This copies contents, formatting, and retains formulas.
        destSheet.Copy(sourceSheet);

        // Save the destination workbook with the transferred worksheet
        destWorkbook.Save("TransferredWorksheet.xlsx");
    }
}
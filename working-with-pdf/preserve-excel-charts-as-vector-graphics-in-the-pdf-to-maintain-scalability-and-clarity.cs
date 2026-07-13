using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the Excel workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // NOTE: Aspose.Cells may provide a property such as `VectorizeCharts` to keep charts as vector graphics.
        // This property is not documented in the supplied reference, so it is left as a placeholder.
        // Uncomment and set the property when using a version that supports it.
        // pdfOptions.VectorizeCharts = true; // <-- requires API evidence

        // Save the workbook as a PDF file
        workbook.Save("output.pdf", pdfOptions);
    }
}
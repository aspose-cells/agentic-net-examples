using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsComponent
{
    // Example component that uses Aspose.Cells resources and implements the .NET dispose pattern
    public class ExcelProcessor : IDisposable
    {
        // Aspose.Cells objects that hold unmanaged resources
        private Workbook _workbook;
        private Worksheet _worksheet;
        private SheetRender _sheetRender;

        // Flag to detect redundant calls
        private bool _disposed = false;

        // Constructor creates and initializes the Aspose.Cells objects
        public ExcelProcessor()
        {
            // Create a new workbook
            _workbook = new Workbook();

            // Access the first worksheet
            _worksheet = _workbook.Worksheets[0];

            // Prepare a SheetRender for rendering the worksheet (example usage)
            _sheetRender = new SheetRender(_worksheet, new ImageOrPrintOptions());
        }

        // Public method demonstrating some operation with the workbook
        public void Process()
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(ExcelProcessor));

            // Example: set a value in cell A1
            _worksheet.Cells["A1"].PutValue("Hello Aspose.Cells!");

            // Example: render the worksheet to an image (not saved, just to illustrate usage)
            _sheetRender.ToImage(0, "output.png");
        }

        // Implementation of IDisposable
        public void Dispose()
        {
            Dispose(true);
            // Suppress finalization since resources have been released
            GC.SuppressFinalize(this);
        }

        // Protected dispose method follows the standard pattern
        protected virtual void Dispose(bool disposing)
        {
            if (_disposed)
                return;

            if (disposing)
            {
                // Dispose managed resources that implement IDisposable
                // Each Aspose.Cells class provides a Dispose method
                _sheetRender?.Dispose();
                _worksheet?.Dispose();
                _workbook?.Dispose();
            }

            // Release unmanaged resources here if any (none in this example)

            _disposed = true;
        }

        // Finalizer to catch cases where Dispose was not called
        ~ExcelProcessor()
        {
            Dispose(false);
        }
    }

    // Example usage of the ExcelProcessor component
    class Program
    {
        static void Main()
        {
            // Use the component within a using block to ensure deterministic disposal
            using (var processor = new ExcelProcessor())
            {
                processor.Process();
                // Additional operations can be performed here
            }

            // At this point, all Aspose.Cells resources have been properly released
        }
    }
}
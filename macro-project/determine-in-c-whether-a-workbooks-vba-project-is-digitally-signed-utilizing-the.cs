using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaSignatureCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            string workbookPath = "sample.xlsm";
            Workbook workbook = new Workbook(workbookPath);
            VbaProject vbaProject = workbook.VbaProject;
            bool isVbaSigned = vbaProject != null && vbaProject.IsSigned;
            int numericResult = isVbaSigned ? 1 : 0;
            Console.WriteLine("VBA Project Signed (numeric): " + numericResult);
        }
    }
}
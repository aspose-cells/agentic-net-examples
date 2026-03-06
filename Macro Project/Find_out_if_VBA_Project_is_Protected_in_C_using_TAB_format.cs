using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaProtectionCheck
{
	public class Program
	{
		public static void Main(string[] args)
		{
			// Path to the Excel file (macro-enabled workbook)
			string filePath = "sample.xlsm";

			// Load the workbook
			Workbook workbook = new Workbook(filePath);

			// Access the VBA project associated with the workbook
			VbaProject vbaProject = workbook.VbaProject;

			// Check if the VBA project is protected
			bool isProtected = vbaProject.IsProtected;

			// Output the result
			Console.WriteLine("Is VBA Project Protected: " + isProtected);
		}
	}
}
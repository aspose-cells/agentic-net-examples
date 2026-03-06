using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaCheck
{
	class Program
	{
		static void Main(string[] args)
		{
			// Path to the macro-enabled Excel file
			string filePath = "sample.xlsm";

			// Load the workbook
			Workbook wb = new Workbook(filePath);

			// Access the VBA project
			VbaProject vbaProject = wb.VbaProject;

			// Determine if the VBA project is protected
			bool isProtected = vbaProject.IsProtected;

			// Determine if the VBA project is locked for viewing
			bool isLockedForViewing = vbaProject.IslockedForViewing;

			// Output the results
			Console.WriteLine("Is VBA Project Protected: " + isProtected);
			Console.WriteLine("Is VBA Project Locked for Viewing: " + isLockedForViewing);
		}
	}
}
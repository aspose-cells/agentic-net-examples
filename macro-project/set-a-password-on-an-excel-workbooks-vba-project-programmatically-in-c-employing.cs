using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace Example
{
	class SetVbaPassword
	{
		static void Main()
		{
			// Create a new workbook
			Workbook workbook = new Workbook();

			// Protect the VBA project with a password (not locked for viewing)
			workbook.VbaProject.Protect(false, "MyVbaPassword");

			// Save the workbook as a macro‑enabled file
			workbook.Save("VbaProtected.xlsm", SaveFormat.Xlsm);
		}
	}
}
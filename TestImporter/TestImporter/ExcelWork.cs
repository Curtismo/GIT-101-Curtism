using System;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Data;

namespace TestImporter
{
	internal class ExcelWork
	{
		private static string GetExcelValue(int x, int y, Application excel, string path)
		{
			Workbook wb = excel.Workbooks.Open(path);
			Worksheet excelSheet = wb.ActiveSheet;
			string test;
			try
			{
				test = excelSheet.Cells[x, y].Value.ToString();
			}
			catch
			{
				test = "";
			}
			wb.Close();
			return test;
		}

		internal static string GetTitle(int currentRow, Microsoft.Office.Interop.Excel.Application excel, string path)
		{
			string title = GetExcelValue(currentRow, 4, excel, path);
			return title;
		}

		internal static string GetDescrition(int currentRow, Microsoft.Office.Interop.Excel.Application excel, string path)
		{
			throw new NotImplementedException();
		}

		internal static string GetParamFileName(int currentRow, Microsoft.Office.Interop.Excel.Application excel, string path)
		{
			throw new NotImplementedException();
		}

		internal static string GetAutomationValue(int currentRow, Microsoft.Office.Interop.Excel.Application excel, string path)
		{
			throw new NotImplementedException();
		}

		internal static string GetParamSheetName(int currentRow, Microsoft.Office.Interop.Excel.Application excel, string path)
		{
			throw new NotImplementedException();
		}

		internal static List<string> GetSteps(int currentRow, Microsoft.Office.Interop.Excel.Application excel, string path)
		{
			throw new NotImplementedException();
		}

		internal static List<string> GetResults(int currentRow, int v, Microsoft.Office.Interop.Excel.Application excel, string path)
		{
			throw new NotImplementedException();
		}

		internal static int GetParentId(int currentRow, Application excel, string path)
		{
			throw new NotImplementedException();
		}

		internal static System.Data.DataTable GetParams(Application excel, string paramPath, string v)
		{
			throw new NotImplementedException();
		}

		internal static void GetStepsForE2E(int v, Application excel, object paramSheet)
		{
			throw new NotImplementedException();
		}
	}
}
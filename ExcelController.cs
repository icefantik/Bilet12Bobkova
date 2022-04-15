using System;
using System.Collections.Generic;
using System.IO;

using Excel = Microsoft.Office.Interop.Excel;

namespace Statement
{
	public class ExcelController : IDisposable
	{
		private Excel.Application _excel;
		private Excel.Workbook _workbook;
		private string _filePath;

		public ExcelController()
		{
			_excel = new Excel.Application();

			/*listRows = new FileParser(System.IO.Path.Combine(Environment.CurrentDirectory, "file.txt")).Parse();

			if (Open(filePath: System.IO.Path.Combine(Environment.CurrentDirectory, "test.xlsx")))
			{
				Clear();

				Set(column: "A", row: 1, data: "Student's Book Number");

				for (int i = 0; i < listRows.Count; i++)
				{
					Set(column: "A", row: i + 2, data: listRows[i].StudentBookNumber);
				}

				Save();
			}*/
		}

		public bool Open(string filePath)
		{
			try
			{
				if (File.Exists(filePath))
				{
					_workbook = _excel.Workbooks.Open(filePath);
				}
				else
				{
					_workbook = _excel.Workbooks.Add();
					_filePath = filePath;
				}
			}
			catch (Exception)
			{
				return false;
			}

			return true;
		}

		public bool Set(string column, int row, object data)
		{
			try
			{
				((Excel.Worksheet)_excel.ActiveSheet).Cells[row, column] = data;
			}
			catch (Exception)
			{
				return false;
			}

			return true;
		}

		public void Clear()
		{
			((Excel.Worksheet)_excel.ActiveSheet).Cells.Clear();
		}

		public void Save()
		{
			if (!string.IsNullOrEmpty(_filePath))
			{
				_workbook.SaveAs(_filePath);

				_filePath = null;
			}
			else
			{
				_workbook.Save();
			}
		}

		public void Dispose()
		{
			try
			{
				_workbook.Close();
			}
			catch (Exception)
			{

			}
		}
	}
}

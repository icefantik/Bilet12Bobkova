using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Statement
{
	/// <summary>
	/// Логика взаимодействия для MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		/*private ExcelController _excel;*/

		private List<StatementRow> _listRows;

		private const string Path = "file.txt";

		public MainWindow()
		{
			InitializeComponent();

			string path = System.IO.Path.Combine(Environment.CurrentDirectory, "file.txt");

			_listRows = new FileParser(path).Parse();

			var Rows = new ObservableCollection<StatementRow>(_listRows);

			JournalRowsGrid.ItemsSource = Rows;

			/*using (var excel = new ExcelController(_listRows)) { }*/

			/*using (var excel = new ExcelController())
			{
				if (excel.Open(filePath: System.IO.Path.Combine(Environment.CurrentDirectory, "test.xlsx")))
				{
					excel.Clear();

					excel.Set(column: "A", row: 1, data: "Student's Book Number");

					for (int i = 0; i < listRows.Count; i++)
					{
						excel.Set(column: "A", row: i + 2, data: listRows[i].StudentBookNumber);
					}

					excel.Save();
				}
			}*/

/*
			try
			{
				using (_excel = new ExcelController(listRows)) { }
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}*/
		}

		private void EditButton_Click(object sender, RoutedEventArgs e)
		{
			var process = System.Diagnostics.Process.Start(System.IO.Path.Combine(Environment.CurrentDirectory, "file.txt"));

			process.WaitForExit();

			string path = System.IO.Path.Combine(Environment.CurrentDirectory, "file.txt");

			var listRows = new FileParser(path).Parse();

			var Rows = new ObservableCollection<StatementRow>(listRows);

			JournalRowsGrid.ItemsSource = Rows;
		}

		private void OpenExcelButton_Click(object sender, RoutedEventArgs e)
		{
			_listRows = new FileParser(Path).Parse();

			using (var excel = new ExcelController())
			{
				if (excel.Open(filePath: System.IO.Path.Combine(Environment.CurrentDirectory, "test.xlsx")))
				{
					excel.Clear();

					excel.Set(column: "A", row: 1, data: "Student's Book Number");

					for (int i = 0; i < _listRows.Count; i++)
					{
						excel.Set(column: "A", row: i + 2, data: _listRows[i].StudentBookNumber);
					}

					excel.Save();
				}
			}
			
			var process = System.Diagnostics.Process.Start(System.IO.Path.Combine(Environment.CurrentDirectory, "test.xlsx"));
		}
	}
}

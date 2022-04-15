using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Statement
{
	public class FileParser
	{

		private string _filePath;

		private string _fileText;

		public string FileText { get => _fileText; }

		public FileParser(string path)
		{
			_filePath = path;

			_fileText = File.ReadAllText(_filePath);

			PrepareFileText();
		}

		public List<StatementRow> Parse()
		{
			List<StatementRow> rows = new List<StatementRow>();

			List<string> stringCells = new List<string>();

			string cell = "";

			string text = _fileText;

			for (int i = 0; i < text.Length; i++)
			{
				if (text[i] == '\r' ||
				(text[i] == '\t' && text[i + 1] != '\t'))
				{
					stringCells.Add(cell);
					cell = "";

					if (text[i] == '\r')
					{
						rows.Add(FillJournalRow(stringCells));

						stringCells = new List<string>();
					}
				}
				else if (text[i] == '\n' || text[i] == '\t')
				{
					continue;
				}
				else
				{
					cell += text[i];
				}
			}

			return rows;
		}

		private void PrepareFileText()
		{
			_fileText = _fileText.Substring(_fileText.IndexOf("\r\n\r\n") + 4);

			if (_fileText.Last() == '\n')
			{
				while (_fileText.Last() == '\n' || _fileText.Last() == '\r')
				{
					_fileText = _fileText.Remove(_fileText.Length - 1);
				}
			}

			_fileText += "\r\n";
		}

		private StatementRow FillJournalRow(List<string> data)
		{
			return new StatementRow()
			{
				Date = DateTime.Parse(data[0]),
				StudentBookNumber = int.Parse(data[1]),
				StudentName = data[2],
				Group = data[3],
				Discipline = data[4],
				Mark = int.Parse(data[5])
			};
		}
	}
}

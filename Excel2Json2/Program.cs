using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Excel2Json2
{
	class Program
	{
		static void Main(string[] args)
		{
			if (args.Length == 0)
			{
				Console.WriteLine("Usage : <programme> <excel file>");
				return;
			}

			var excelFileName = args[0];
			var directory = Path.GetDirectoryName(excelFileName);
			var jsonFileName = Path.Combine(directory, Path.GetFileNameWithoutExtension(excelFileName));
			var dictData = new Dictionary<string, Data>()
			{
				{ "ENG", new Data() },
				{ "VIE", new Data() },
				{ "GER", new Data() },
			};

			var sheetNames = new[] { "Normal", "Hard" };

			var sheets = Excel.Workbook.Worksheets(excelFileName).Where(x=>x.Rows.Length>0).ToList();

			for(int i = 0; i < sheets.Count; i++)
			{
				foreach (var item in dictData)
				{
					item.Value.Add(new Level(sheetNames[i]));
				}

				var sheet = sheets[i];
				int rowIndex = 0;
				var lastAnswer = "";

				foreach (var row in sheet.Rows.Skip(1))
				{
					rowIndex++;
					var language = row.Cells[1].Text;
					var question = row.Cells[2].Text;
					var answer = row.Cells[3].Text;

					if (string.IsNullOrEmpty(answer))
					{
						answer = lastAnswer;
					}
					else
					{
						lastAnswer = answer;
					}
					if (!string.IsNullOrEmpty(question))
					{
						dictData[language][i].Questions.Add(new Question()
						{
							Answer = answer,
							Title = question
						});
					}
				}
			}

			foreach (var item in dictData)
			{
				var fileName = string.Format("{0}_{1}.json", jsonFileName, item.Key);
				File.WriteAllText(fileName, JsonConvert.SerializeObject(item.Value));
			}
		}
	}
}

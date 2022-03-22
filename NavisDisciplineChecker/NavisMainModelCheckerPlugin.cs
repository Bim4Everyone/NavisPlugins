using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;
using Autodesk.Navisworks.Api.DocumentParts;
using Autodesk.Navisworks.Api.Clash;
using Autodesk.Navisworks.Api.Interop;

using DevExpress.Spreadsheet;
using DevExpress.XtraSpreadsheet.Model;

using Row = DevExpress.Spreadsheet.Row;
using Worksheet = DevExpress.Spreadsheet.Worksheet;

namespace NavisDisciplineChecker {
    [Plugin("NavisMainModelCheckerPlugin", "budaevaler",
        DisplayName = "Создание сводной модели", ToolTip = "Создает сводную модель")]
    public class NavisMainModelCheckerPlugin : AddInPlugin {
        public override int Execute(params string[] parameters) {
            var document = Application.ActiveDocument;

            var nwfFilePath = document.FileName;
            var rootPath = Path.GetDirectoryName(nwfFilePath);

            var nwdFileName
                = DateTime.Now.ToString("yy.MM.dd") + "_" +
                  Path.ChangeExtension(Path.GetFileName(nwfFilePath), ".nwd");
            
            DocumentClash clash = document.GetClash();
            clash.TestsData.TestsRunAllTests();

            document.SaveFile(nwfFilePath);
            document.SaveFile(Path.Combine(rootPath, "!Отчёты", nwdFileName));

            var result = clash.TestsData.Tests
                .OfType<ClashTest>()
                .Select(item => new {
                    TestName = item.DisplayName,
                    Comments = item.Children
                        .OfType<ClashResult>()
                        .Where(comment => comment.Status == ClashResultStatus.New
                                          || comment.Status == ClashResultStatus.Active)
                        .ToList()
                })
                .Select(item => new {TestName = item.TestName, Count = item.Comments.Count})
                .ToDictionary(item => item.TestName, item => item.Count);

            var reportNamePath = Directory.GetFiles(rootPath, "*.xlsx")
                .FirstOrDefault();

            using(Workbook workbook = new Workbook()) {
                if(string.IsNullOrEmpty(reportNamePath)) {
                    reportNamePath = Path.Combine(rootPath, "Прогресс устранения коллизий.xlsx");
                } else {
                    workbook.LoadDocument(reportNamePath);
                }

                Worksheet worksheet = workbook.Worksheets.FirstOrDefault();
                int lastColumnIndex = worksheet.Columns.LastUsedIndex;
                var cell = worksheet.Cells[0, lastColumnIndex];
                if((DateTime.Now.Date - cell.Value.DateTimeValue.Date) > TimeSpan.FromDays(1)) {
                    lastColumnIndex++;
                }

                worksheet.Cells[0, lastColumnIndex].NumberFormat = "yy.MM.dd";
                worksheet.Cells[0, lastColumnIndex].SetValueFromText(DateTime.Now.ToString("dd.MM.yy"));
                for(int index = 1; index <= worksheet.Rows.LastUsedIndex; index++) {
                    var testName = worksheet.Cells[index, 0].DisplayText;
                    if(result.TryGetValue(testName, out int count)) {
                        result.Remove(testName);
                        worksheet.Cells[index, lastColumnIndex].SetValueFromText(count.ToString());
                    } else {
                        worksheet.Cells[index, lastColumnIndex].SetValueFromText(null);
                    }
                }

                foreach(KeyValuePair<string, int> kvp in result) {
                    int rowIndex = worksheet.Rows.LastUsedIndex + 1;
                    worksheet.Cells[rowIndex, 0]
                        .SetValueFromText(kvp.Key);
                    
                    worksheet.Cells[rowIndex, lastColumnIndex]
                        .SetValueFromText(kvp.Value.ToString());
                }

                workbook.SaveDocument(reportNamePath);
            }

            Process.Start(reportNamePath);
            return 0;
        }
    }
}
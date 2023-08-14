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

using NavisDisciplineChecker.Tables;

namespace NavisDisciplineChecker {
    [Plugin("NavisMainModelCheckerPlugin", "budaevaler",
        DisplayName = "Создание сводной модели", ToolTip = "Создает сводную модель")]
    public class NavisMainModelCheckerPlugin : AddInPlugin {
        public override int Execute(params string[] parameters) {
            var document = Application.ActiveDocument;

            var nwfFilePath = document.FileName;
            var logFileName = Path.ChangeExtension(nwfFilePath, ".log");
            var errorFileName = Path.ChangeExtension(nwfFilePath, ".error");

            var rootPath = Path.GetDirectoryName(nwfFilePath);
            var projectName = Path.GetFileName(nwfFilePath).Split('_').FirstOrDefault();

            var currentDate = parameters.FirstOrDefault()
                              ?? DateTime.Now.ToString("yy.MM.dd");

            var nwdFileName
                = currentDate + "_" +
                  Path.ChangeExtension(Path.GetFileName(nwfFilePath), ".nwd");


            var nwdFilePath = Path.Combine(rootPath, "!Отчёты", currentDate, nwdFileName);
            Directory.CreateDirectory(Path.GetDirectoryName(nwdFilePath));

            try {
                using(SimpleLogger logger = new SimpleLogger(logFileName)) {
                    DocumentClash clash = document.GetClash();
                    if(clash.TestsData.Tests.Count == 0) {
                        throw new InvalidOperationException(
                            $"У файла {Path.GetFileName(nwfFilePath)} не настроены правила коллизий.");
                    }

                    clash.TestsData.TestsRunAllTests();
                    logger.WriteLine("Запуск поиска коллизий.");

                    document.SaveFile(nwfFilePath);
                    logger.WriteLine($"Сохранение файла NWF \"{nwfFilePath}\".");

                    document.SaveFile(nwdFilePath);
                    logger.WriteLine($"Сохранение файла NWD \"{nwdFilePath}\".");

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

                    var reportNamePath =
                        Path.Combine(rootPath, projectName + "_" + "Прогресс устранения коллизий.csv");

                    var worksheet = new CsvTable();
                    if(File.Exists(reportNamePath)) {
                        worksheet.LoadDocument(reportNamePath);
                        logger.WriteLine($"Открытие файла csv \"{reportNamePath}\".");
                    }
                    
                    Column column = worksheet.CreateColumn(currentDate);
                    for(int index = 0; index < worksheet.Rows.Count; index++) {
                        var testName = worksheet.Rows[index][0];
                        if(result.TryGetValue(testName, out int count)) {
                            result.Remove(testName);
                            worksheet.Rows[index][column] = count.ToString();
                        }
                    }

                    foreach(KeyValuePair<string, int> kvp in result) {
                        var row = worksheet.CreateRow();
                        row[0] = kvp.Key;
                        row[column] = kvp.Value.ToString();
                    }

                    worksheet.SaveDocument(reportNamePath);
                    logger.WriteLine($"Сохранение файла csv \"{reportNamePath}\".");
                }

                return 0;
            } catch(Exception ex) {
                File.WriteAllText(errorFileName, ex.ToString());
                return 1;
            }
        }
    }
}
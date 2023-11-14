using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;

namespace NavisDisciplineChecker {

    [Plugin("NavisDisciplineChecker", "budaevaler",
        DisplayName = "Проверка nwf", ToolTip = "Проверка nwf по разделам")]
    public class NavisDisciplineCheckerPlugin : AddInPlugin {
        public override int Execute(params string[] parameters) {
            return CombineDisciplineFiles();
        }

        private int CombineDisciplineFiles() {
            var document = Application.ActiveDocument;

            var nwfFilePath = document.FileName;

            var logFileName = Path.ChangeExtension(nwfFilePath, ".log");
            var errorFileName = Path.ChangeExtension(nwfFilePath, ".error");

            var rootPath = Path.GetDirectoryName(nwfFilePath);
            var nwdRootPath = Path.GetDirectoryName(rootPath);

            string nwdFilePath =
                Path.Combine(nwdRootPath, Path.ChangeExtension(Path.GetFileName(nwfFilePath), ".nwd"));

            string mainModelPath =
                Path.Combine(Path.GetDirectoryName(nwdRootPath), "Сводная модель");

            string nwdMainModelFilePath =
                Path.Combine(mainModelPath, Path.ChangeExtension(Path.GetFileName(nwfFilePath), ".nwd"));

            var nwcFileNames =
                Directory.GetFiles(rootPath, "*.nwc");

            try {
                using(SimpleLogger logger = new SimpleLogger(logFileName)) {
                    document.Clear();
                    logger.WriteLine($"Очистка файла NWF \"{nwfFilePath}\".");

                    document.AppendFiles(nwcFileNames.Select(item => item).OrderBy(item => item));
                    logger.WriteLine($"Добавление файлов NWС " +
                                     $"{Environment.NewLine} - " +
                                     $"{string.Join(Environment.NewLine + " - ", nwcFileNames.Select(item => item).OrderBy(item => item))}");

                    document.SavedViewpoints.Clear();
                    logger.WriteLine($"Очистка точек обзора.");

                    document.SaveFile(nwfFilePath);
                    logger.WriteLine($"Сохранение файла NWF \"{nwfFilePath}\".");

                    document.SaveFile(nwdFilePath);
                    logger.WriteLine($"Сохранение файла NWD \"{nwdFilePath}\".");

                    File.Copy(nwdFilePath, nwdMainModelFilePath, true);
                    logger.WriteLine($"Копирование файла NWD в сводную модель \"{nwdMainModelFilePath}\".");
                    
                    return 0;
                }
            } catch(Exception ex) {
                File.WriteAllText(errorFileName, ex.ToString());
                return 1;
            }
        }
    }
}
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;

namespace NavisDisciplineChecker {

    [Plugin("NavisDisciplineChecker", "budaevaler", DisplayName = "Проверка nwf", ToolTip = "Проверка nwf по разделам")]
    public class NavisDisciplineCheckerPlugin : AddInPlugin {
        public override int Execute(params string[] parameters) {
            CombineDisciplineFiles();
            return 0;
        }

        private void CombineDisciplineFiles() {
            var document = Application.ActiveDocument;

            var nwfFilePath = document.FileName;
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

            document.Clear();
            document.AppendFiles(nwcFileNames.Select(item => item).OrderBy(item => item));

            document.SavedViewpoints.Clear();
            document.SaveFile(nwfFilePath);
            document.SaveFile(nwdFilePath);

            File.Copy(nwdFilePath, nwdMainModelFilePath, true);
        }
    }
}
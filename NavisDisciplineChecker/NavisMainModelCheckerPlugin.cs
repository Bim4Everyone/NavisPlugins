using System;
using System.IO;
using System.Linq;

using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;

namespace NavisDisciplineChecker
{
    [Plugin("NavisMainModelCheckerPlugin", "budaevaler", DisplayName = "Создание сводной модели", ToolTip = "Создает сводную модель")]
    public class NavisMainModelCheckerPlugin : AddInPlugin {
        public override int Execute(params string[] parameters) {
            var document = Application.ActiveDocument;
            
            var nwfFilePath = document.FileName;
            var rootPath = Path.GetDirectoryName(nwfFilePath);

            var nwdFileName 
                = DateTime.Now.ToString("yy.mm.dd") + "_" +
                  Path.ChangeExtension(Path.GetFileName(nwfFilePath), ".nwd");
            
            var nwdFileNames =
                Directory.GetFiles(rootPath, "*.nwd");
            
            document.Clear();
            document.AppendFiles(nwdFileNames.Select(item => item).OrderBy(item => item));

            document.SavedViewpoints.Clear();
            document.SaveFile(nwfFilePath);
            document.SaveFile(Path.Combine(rootPath, "!Отчёты", nwdFileName));

            return 0;
        }
    }
}
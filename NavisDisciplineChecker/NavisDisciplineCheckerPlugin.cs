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
            OverwriteModels();
            return 0;
        }

        private void OverwriteModels() {
            var document = Application.ActiveDocument;
            var path = Path.GetDirectoryName(document.FileName);
            var nwcFiles = Directory.GetFiles(path, "*.nwc");
            document.Clear();
            foreach(var nwc in nwcFiles) {
                document.AppendFile(nwc);
            }
            //document.SaveFile(document.FileName);
        }

        private void UniteFiles() {
            var document = Application.ActiveDocument;
            var path = Path.GetDirectoryName(document.FileName);
            var nwcFileNames = Directory.GetFiles(path, "*.nwc")
                .Select(item => Path.GetFileName(item))
                .GroupBy(item => GetDiscipline(item));

            foreach(IGrouping<string, string> nwcFileName in nwcFileNames) {
                //string newNwfFileName = 
            }
        }

        private string GetDiscipline(string nwcName) {
            return nwcName.Split('_').ElementAtOrDefault(2);
        }

    }


    internal class NwcGroup {
        public string Discipline { get; set; }
        public List<string> NwcFiles { get; set; }
    }
}

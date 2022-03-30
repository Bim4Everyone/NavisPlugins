using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NavisDisciplineChecker.ClashModel {
    public class Clash {
        public Clash(Autodesk.Navisworks.Api.Clash.ClashResult clashResult) {
            Name = clashResult.DisplayName;
            ClashElement1 = new ClashElement(clashResult.Item1);
            ClashElement2 = new ClashElement(clashResult.Item2);
        }

        public string Name { get; set; }
        public ClashElement ClashElement1 { get; set; }
        public ClashElement ClashElement2 { get; set; }
    }
}

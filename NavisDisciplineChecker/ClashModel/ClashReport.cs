using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NavisDisciplineChecker.ClashModel {
    public class ClashReport {
        public string Name { get; set; }
        public List<Clash> Clashes { get; set; } = new List<Clash>();
    }
}

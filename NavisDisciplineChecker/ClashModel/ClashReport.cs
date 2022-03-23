using System.Collections.Generic;

namespace NavisDisciplineChecker
{
    public class ClashReport {
        public string Name { get; set; }
        public List<Clash> Clashes { get; set; } = new List<Clash>();
    }
}
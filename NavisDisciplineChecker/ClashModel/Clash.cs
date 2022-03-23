using Autodesk.Navisworks.Api.Clash;

namespace NavisDisciplineChecker
{
    public class Clash {
        public Clash(ClashResult clashResult) {
            Name = clashResult.DisplayName;
            ClashElement1 = new ClashElement(clashResult.Item1);
            ClashElement2 = new ClashElement(clashResult.Item2);
        }

        public string Name { get; set; }
        public ClashElement ClashElement1 { get; set; }
        public ClashElement ClashElement2 { get; set; }
    }
}
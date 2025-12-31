namespace ValveFlangeMulti.Models
{
    public sealed class PmsRow
    {
        public int RowIndex { get; set; }

        public string Class { get; set; } = "";        // A
        public string Alt { get; set; } = "";          // G
        public string ItemType { get; set; } = "";     // H
        public string ItemName { get; set; } = "";     // I
        public string ConnectionType { get; set; } = "";// L
        public string FamilyName { get; set; } = "";   // M
        public string TypeName { get; set; } = "";     // N

        public double MainFrom { get; set; }           // C
        public double MainTo { get; set; }             // D
    }
}

namespace MIR.Models
{
    public class Material
    {
        public Guid Id { get; set; } = Guid.NewGuid();
        public string Code { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public string Unit { get; set; } = "pcs";
        public decimal CurrentStock { get; set; } = 0;
        public decimal MinStockLevel { get; set; } = 0;
        public decimal UnitPrice { get; set; } = 0;
        public string Category { get; set; } = "Raw Materials";
        public DateTime LastUpdated { get; set; } = DateTime.Now;

        public bool IsLowStock => CurrentStock <= MinStockLevel;

        public Material() { }

        public Material(string code, string name, string unit, string category)
        {
            Code = code;
            Name = name;
            Unit = unit;
            Category = category;
        }
    }

    public static class MaterialCategories
    {
        public static readonly string[] Categories = new[]
        {
            "Raw Materials",
            "Electronics",
            "Packaging",
            "Hardware",
            "Chemicals",
            "Other"
        };
    }
}

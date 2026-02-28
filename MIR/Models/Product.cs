namespace MIR.Models
{
    public class Product
    {
        public Guid Id { get; set; } = Guid.NewGuid();
        public string Code { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public string Category { get; set; } = "Finished Goods";
        public bool IsActive { get; set; } = true;
        public DateTime CreatedAt { get; set; } = DateTime.Now;

        public Product() { }

        public Product(string code, string name, string category)
        {
            Code = code;
            Name = name;
            Category = category;
        }
    }

    public static class ProductCategories
    {
        public static readonly string[] Categories = new[]
        {
            "Finished Goods",
            "Semi-Finished",
            "Assemblies",
            "Other"
        };
    }
}

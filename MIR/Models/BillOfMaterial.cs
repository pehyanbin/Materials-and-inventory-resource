namespace MIR.Models
{
    public class BillOfMaterial
    {
        public Guid Id { get; set; } = Guid.NewGuid();
        public Guid ProductId { get; set; }
        public Guid MaterialId { get; set; }
        public decimal Quantity { get; set; } = 1;

        // Navigation properties (populated at runtime)
        public string? ProductName { get; set; }
        public string? ProductCode { get; set; }
        public string? MaterialName { get; set; }
        public string? MaterialCode { get; set; }
        public string? MaterialUnit { get; set; }

        public BillOfMaterial() { }

        public BillOfMaterial(Guid productId, Guid materialId, decimal quantity)
        {
            ProductId = productId;
            MaterialId = materialId;
            Quantity = quantity;
        }
    }
}

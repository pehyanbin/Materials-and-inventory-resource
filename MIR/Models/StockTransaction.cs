namespace MIR.Models
{
    public enum TransactionType
    {
        IN,
        OUT
    }

    public class StockTransaction
    {
        public Guid Id { get; set; } = Guid.NewGuid();
        public Guid MaterialId { get; set; }
        public TransactionType Type { get; set; }
        public decimal Quantity { get; set; }
        public string Notes { get; set; } = string.Empty;
        public Guid UserId { get; set; }
        public DateTime TransactionDate { get; set; } = DateTime.Now;

        // Navigation properties (populated at runtime)
        public string? MaterialName { get; set; }
        public string? MaterialCode { get; set; }
        public string? UserName { get; set; }

        public StockTransaction() { }

        public StockTransaction(Guid materialId, TransactionType type, decimal quantity, Guid userId, string notes = "")
        {
            MaterialId = materialId;
            Type = type;
            Quantity = quantity;
            UserId = userId;
            Notes = notes;
        }
    }
}

using MIR.Models;

namespace MIR.Services
{
    public interface IExcelService
    {
        // Users
        List<User> GetAllUsers();
        User? GetUserById(Guid id);
        User? GetUserByUsername(string username);
        void AddUser(User user);
        void UpdateUser(User user);
        void DeleteUser(Guid id);

        // Materials
        List<Material> GetAllMaterials();
        Material? GetMaterialById(Guid id);
        void AddMaterial(Material material);
        void UpdateMaterial(Material material);
        void DeleteMaterial(Guid id);
        List<Material> GetLowStockMaterials();

        // Products
        List<Product> GetAllProducts();
        Product? GetProductById(Guid id);
        void AddProduct(Product product);
        void UpdateProduct(Product product);
        void DeleteProduct(Guid id);

        // Bill of Materials
        List<BillOfMaterial> GetAllBOM();
        List<BillOfMaterial> GetBOMByProductId(Guid productId);
        void AddBOM(BillOfMaterial bom);
        void UpdateBOM(BillOfMaterial bom);
        void DeleteBOM(Guid id);
        void DeleteBOMByProductId(Guid productId);

        // Stock Transactions
        List<StockTransaction> GetAllTransactions();
        List<StockTransaction> GetTransactionsByMaterialId(Guid materialId);
        void AddTransaction(StockTransaction transaction);

        // Stock Check
        bool CanProduce(Guid productId, int quantity, out List<(Material material, decimal required, decimal available, decimal shortage)> shortages);
        int MaxProducible(Guid productId);
        List<(Material Material, decimal Required, decimal Available, decimal Shortage, decimal ImpactPercentage)> GetProductionAnalysis(Guid productId, int quantity);

        // Import/Export
        void ExportToExcel(string filePath);
        void ImportFromExcel(string filePath);

        // Initialize
        void Initialize();

        // Settings
        string GetSetting(string key, string defaultValue);
        void UpdateSetting(string key, string value);
    }
}

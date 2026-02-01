using MIR.Models;
using OfficeOpenXml;
using System.IO;

namespace MIR.Services
{
    public class ExcelService : IExcelService
    {
        private readonly string _dataFilePath;
        private readonly object _lock = new();

        public ExcelService()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
            // Store data in the same directory as the executable
            var exePath = AppDomain.CurrentDomain.BaseDirectory;
            _dataFilePath = Path.Combine(exePath, "data.xlsx");
            
            /*
            // Store data in AppData\MIR folder
            var appDataPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "MIR"
            );
            
            if (!Directory.Exists(appDataPath))
            {
                Directory.CreateDirectory(appDataPath);
            }
            
            _dataFilePath = Path.Combine(appDataPath, "data.xlsx");
            */
        }

        public void Initialize()
        {
            lock (_lock)
            {
                if (!File.Exists(_dataFilePath))
                {
                    CreateNewDataFile();
                }
            }
        }

        private void CreateNewDataFile()
        {
            using var package = new ExcelPackage();
            
            // Users sheet
            var usersSheet = package.Workbook.Worksheets.Add("Users");
            usersSheet.Cells[1, 1].Value = "Id";
            usersSheet.Cells[1, 2].Value = "Username";
            usersSheet.Cells[1, 3].Value = "PasswordHash";
            usersSheet.Cells[1, 4].Value = "FullName";
            usersSheet.Cells[1, 5].Value = "Role";
            usersSheet.Cells[1, 6].Value = "IsActive";
            usersSheet.Cells[1, 7].Value = "CreatedAt";
            
            // Add default admin user
            var adminHash = BCrypt.Net.BCrypt.HashPassword("admin123");
            usersSheet.Cells[2, 1].Value = Guid.NewGuid().ToString();
            usersSheet.Cells[2, 2].Value = "admin";
            usersSheet.Cells[2, 3].Value = adminHash;
            usersSheet.Cells[2, 4].Value = "Administrator";
            usersSheet.Cells[2, 5].Value = "Admin";
            usersSheet.Cells[2, 6].Value = true;
            usersSheet.Cells[2, 7].Value = DateTime.Now.ToString("O");
            
            // Materials sheet
            var materialsSheet = package.Workbook.Worksheets.Add("Materials");
            materialsSheet.Cells[1, 1].Value = "Id";
            materialsSheet.Cells[1, 2].Value = "Code";
            materialsSheet.Cells[1, 3].Value = "Name";
            materialsSheet.Cells[1, 4].Value = "Description";
            materialsSheet.Cells[1, 5].Value = "Unit";
            materialsSheet.Cells[1, 6].Value = "CurrentStock";
            materialsSheet.Cells[1, 7].Value = "MinStockLevel";
            materialsSheet.Cells[1, 8].Value = "UnitPrice";
            materialsSheet.Cells[1, 9].Value = "Category";
            materialsSheet.Cells[1, 10].Value = "LastUpdated";
            
            // Products sheet
            var productsSheet = package.Workbook.Worksheets.Add("Products");
            productsSheet.Cells[1, 1].Value = "Id";
            productsSheet.Cells[1, 2].Value = "Code";
            productsSheet.Cells[1, 3].Value = "Name";
            productsSheet.Cells[1, 4].Value = "Description";
            productsSheet.Cells[1, 5].Value = "Category";
            productsSheet.Cells[1, 6].Value = "IsActive";
            productsSheet.Cells[1, 7].Value = "CreatedAt";
            
            // BillOfMaterials sheet
            var bomSheet = package.Workbook.Worksheets.Add("BillOfMaterials");
            bomSheet.Cells[1, 1].Value = "Id";
            bomSheet.Cells[1, 2].Value = "ProductId";
            bomSheet.Cells[1, 3].Value = "MaterialId";
            bomSheet.Cells[1, 4].Value = "Quantity";
            
            // StockTransactions sheet
            var transactionsSheet = package.Workbook.Worksheets.Add("StockTransactions");
            transactionsSheet.Cells[1, 1].Value = "Id";
            transactionsSheet.Cells[1, 2].Value = "MaterialId";
            transactionsSheet.Cells[1, 3].Value = "Type";
            transactionsSheet.Cells[1, 4].Value = "Quantity";
            transactionsSheet.Cells[1, 5].Value = "Notes";
            transactionsSheet.Cells[1, 6].Value = "UserId";
            transactionsSheet.Cells[1, 7].Value = "TransactionDate";
            
            package.SaveAs(new FileInfo(_dataFilePath));
        }

        #region Users

        public List<User> GetAllUsers()
        {
            lock (_lock)
            {
                var users = new List<User>();
                using var package = new ExcelPackage(new FileInfo(_dataFilePath));
                var sheet = package.Workbook.Worksheets["Users"];
                
                if (sheet == null) return users;
                
                var rowCount = sheet.Dimension?.Rows ?? 0;
                for (int row = 2; row <= rowCount; row++)
                {
                    if (string.IsNullOrEmpty(sheet.Cells[row, 1].Text)) continue;
                    
                    users.Add(new User
                    {
                        Id = Guid.TryParse(sheet.Cells[row, 1].Text, out var id) ? id : Guid.NewGuid(),
                        Username = sheet.Cells[row, 2].Text,
                        PasswordHash = sheet.Cells[row, 3].Text,
                        FullName = sheet.Cells[row, 4].Text,
                        Role = Enum.TryParse<UserRole>(sheet.Cells[row, 5].Text, out var role) ? role : UserRole.Staff,
                        IsActive = ParseBool(sheet.Cells[row, 6]),
                        CreatedAt = DateTime.TryParse(sheet.Cells[row, 7].Text, out var date) ? date : DateTime.Now
                    });
                }
                
                return users;
            }
        }

        public User? GetUserById(Guid id)
        {
            return GetAllUsers().FirstOrDefault(u => u.Id == id);
        }

        public User? GetUserByUsername(string username)
        {
            return GetAllUsers().FirstOrDefault(u => u.Username.Equals(username, StringComparison.OrdinalIgnoreCase));
        }

        public void AddUser(User user)
        {
            lock (_lock)
            {
                using var package = new ExcelPackage(new FileInfo(_dataFilePath));
                var sheet = package.Workbook.Worksheets["Users"];
                var rowCount = (sheet.Dimension?.Rows ?? 1) + 1;
                
                sheet.Cells[rowCount, 1].Value = user.Id.ToString();
                sheet.Cells[rowCount, 2].Value = user.Username;
                sheet.Cells[rowCount, 3].Value = user.PasswordHash;
                sheet.Cells[rowCount, 4].Value = user.FullName;
                sheet.Cells[rowCount, 5].Value = user.Role.ToString();
                sheet.Cells[rowCount, 6].Value = user.IsActive;
                sheet.Cells[rowCount, 7].Value = user.CreatedAt.ToString("O");
                
                package.Save();
            }
        }

        public void UpdateUser(User user)
        {
            lock (_lock)
            {
                using var package = new ExcelPackage(new FileInfo(_dataFilePath));
                var sheet = package.Workbook.Worksheets["Users"];
                var rowCount = sheet.Dimension?.Rows ?? 0;
                
                for (int row = 2; row <= rowCount; row++)
                {
                    if (sheet.Cells[row, 1].Text == user.Id.ToString())
                    {
                        sheet.Cells[row, 2].Value = user.Username;
                        sheet.Cells[row, 3].Value = user.PasswordHash;
                        sheet.Cells[row, 4].Value = user.FullName;
                        sheet.Cells[row, 5].Value = user.Role.ToString();
                        sheet.Cells[row, 6].Value = user.IsActive;
                        sheet.Cells[row, 7].Value = user.CreatedAt.ToString("O");
                        break;
                    }
                }
                
                package.Save();
            }
        }

        public void DeleteUser(Guid id)
        {
            lock (_lock)
            {
                using var package = new ExcelPackage(new FileInfo(_dataFilePath));
                var sheet = package.Workbook.Worksheets["Users"];
                var rowCount = sheet.Dimension?.Rows ?? 0;
                
                for (int row = 2; row <= rowCount; row++)
                {
                    if (sheet.Cells[row, 1].Text == id.ToString())
                    {
                        sheet.DeleteRow(row);
                        break;
                    }
                }
                
                package.Save();
            }
        }

        #endregion

        #region Materials

        public List<Material> GetAllMaterials()
        {
            lock (_lock)
            {
                var materials = new List<Material>();
                using var package = new ExcelPackage(new FileInfo(_dataFilePath));
                var sheet = package.Workbook.Worksheets["Materials"];
                
                if (sheet == null) return materials;
                
                var rowCount = sheet.Dimension?.Rows ?? 0;
                for (int row = 2; row <= rowCount; row++)
                {
                    if (string.IsNullOrEmpty(sheet.Cells[row, 1].Text)) continue;
                    
                    materials.Add(new Material
                    {
                        Id = Guid.TryParse(sheet.Cells[row, 1].Text, out var id) ? id : Guid.NewGuid(),
                        Code = sheet.Cells[row, 2].Text,
                        Name = sheet.Cells[row, 3].Text,
                        Description = sheet.Cells[row, 4].Text,
                        Unit = sheet.Cells[row, 5].Text,
                        CurrentStock = decimal.TryParse(sheet.Cells[row, 6].Text, out var stock) ? stock : 0,
                        MinStockLevel = decimal.TryParse(sheet.Cells[row, 7].Text, out var min) ? min : 0,
                        UnitPrice = decimal.TryParse(sheet.Cells[row, 8].Text, out var price) ? price : 0,
                        Category = sheet.Cells[row, 9].Text,
                        LastUpdated = DateTime.TryParse(sheet.Cells[row, 10].Text, out var date) ? date : DateTime.Now
                    });
                }
                
                return materials;
            }
        }

        public Material? GetMaterialById(Guid id)
        {
            return GetAllMaterials().FirstOrDefault(m => m.Id == id);
        }

        public void AddMaterial(Material material)
        {
            lock (_lock)
            {
                using var package = new ExcelPackage(new FileInfo(_dataFilePath));
                var sheet = package.Workbook.Worksheets["Materials"];
                var rowCount = (sheet.Dimension?.Rows ?? 1) + 1;
                
                sheet.Cells[rowCount, 1].Value = material.Id.ToString();
                sheet.Cells[rowCount, 2].Value = material.Code;
                sheet.Cells[rowCount, 3].Value = material.Name;
                sheet.Cells[rowCount, 4].Value = material.Description;
                sheet.Cells[rowCount, 5].Value = material.Unit;
                sheet.Cells[rowCount, 6].Value = material.CurrentStock;
                sheet.Cells[rowCount, 7].Value = material.MinStockLevel;
                sheet.Cells[rowCount, 8].Value = material.UnitPrice;
                sheet.Cells[rowCount, 9].Value = material.Category;
                sheet.Cells[rowCount, 10].Value = material.LastUpdated.ToString("O");
                
                package.Save();
            }
        }

        public void UpdateMaterial(Material material)
        {
            lock (_lock)
            {
                using var package = new ExcelPackage(new FileInfo(_dataFilePath));
                var sheet = package.Workbook.Worksheets["Materials"];
                var rowCount = sheet.Dimension?.Rows ?? 0;
                
                for (int row = 2; row <= rowCount; row++)
                {
                    if (sheet.Cells[row, 1].Text == material.Id.ToString())
                    {
                        sheet.Cells[row, 2].Value = material.Code;
                        sheet.Cells[row, 3].Value = material.Name;
                        sheet.Cells[row, 4].Value = material.Description;
                        sheet.Cells[row, 5].Value = material.Unit;
                        sheet.Cells[row, 6].Value = material.CurrentStock;
                        sheet.Cells[row, 7].Value = material.MinStockLevel;
                        sheet.Cells[row, 8].Value = material.UnitPrice;
                        sheet.Cells[row, 9].Value = material.Category;
                        sheet.Cells[row, 10].Value = DateTime.Now.ToString("O");
                        break;
                    }
                }
                
                package.Save();
            }
        }

        public void DeleteMaterial(Guid id)
        {
            lock (_lock)
            {
                using var package = new ExcelPackage(new FileInfo(_dataFilePath));
                var sheet = package.Workbook.Worksheets["Materials"];
                var rowCount = sheet.Dimension?.Rows ?? 0;
                
                for (int row = 2; row <= rowCount; row++)
                {
                    if (sheet.Cells[row, 1].Text == id.ToString())
                    {
                        sheet.DeleteRow(row);
                        break;
                    }
                }
                
                package.Save();
            }
        }

        public List<Material> GetLowStockMaterials()
        {
            return GetAllMaterials().Where(m => m.IsLowStock).ToList();
        }

        #endregion

        #region Products

        public List<Product> GetAllProducts()
        {
            lock (_lock)
            {
                var products = new List<Product>();
                using var package = new ExcelPackage(new FileInfo(_dataFilePath));
                var sheet = package.Workbook.Worksheets["Products"];
                
                if (sheet == null) return products;
                
                var rowCount = sheet.Dimension?.Rows ?? 0;
                for (int row = 2; row <= rowCount; row++)
                {
                    if (string.IsNullOrEmpty(sheet.Cells[row, 1].Text)) continue;
                    
                    products.Add(new Product
                    {
                        Id = Guid.TryParse(sheet.Cells[row, 1].Text, out var id) ? id : Guid.NewGuid(),
                        Code = sheet.Cells[row, 2].Text,
                        Name = sheet.Cells[row, 3].Text,
                        Description = sheet.Cells[row, 4].Text,
                        Category = sheet.Cells[row, 5].Text,
                        IsActive = ParseBool(sheet.Cells[row, 6]),
                        CreatedAt = DateTime.TryParse(sheet.Cells[row, 7].Text, out var date) ? date : DateTime.Now
                    });
                }
                
                return products;
            }
        }

        public Product? GetProductById(Guid id)
        {
            return GetAllProducts().FirstOrDefault(p => p.Id == id);
        }

        public void AddProduct(Product product)
        {
            lock (_lock)
            {
                using var package = new ExcelPackage(new FileInfo(_dataFilePath));
                var sheet = package.Workbook.Worksheets["Products"];
                var rowCount = (sheet.Dimension?.Rows ?? 1) + 1;
                
                sheet.Cells[rowCount, 1].Value = product.Id.ToString();
                sheet.Cells[rowCount, 2].Value = product.Code;
                sheet.Cells[rowCount, 3].Value = product.Name;
                sheet.Cells[rowCount, 4].Value = product.Description;
                sheet.Cells[rowCount, 5].Value = product.Category;
                sheet.Cells[rowCount, 6].Value = product.IsActive;
                sheet.Cells[rowCount, 7].Value = product.CreatedAt.ToString("O");
                
                package.Save();
            }
        }

        public void UpdateProduct(Product product)
        {
            lock (_lock)
            {
                using var package = new ExcelPackage(new FileInfo(_dataFilePath));
                var sheet = package.Workbook.Worksheets["Products"];
                var rowCount = sheet.Dimension?.Rows ?? 0;
                
                for (int row = 2; row <= rowCount; row++)
                {
                    if (sheet.Cells[row, 1].Text == product.Id.ToString())
                    {
                        sheet.Cells[row, 2].Value = product.Code;
                        sheet.Cells[row, 3].Value = product.Name;
                        sheet.Cells[row, 4].Value = product.Description;
                        sheet.Cells[row, 5].Value = product.Category;
                        sheet.Cells[row, 6].Value = product.IsActive;
                        sheet.Cells[row, 7].Value = product.CreatedAt.ToString("O");
                        break;
                    }
                }
                
                package.Save();
            }
        }

        public void DeleteProduct(Guid id)
        {
            lock (_lock)
            {
                using var package = new ExcelPackage(new FileInfo(_dataFilePath));
                var sheet = package.Workbook.Worksheets["Products"];
                var rowCount = sheet.Dimension?.Rows ?? 0;
                
                for (int row = 2; row <= rowCount; row++)
                {
                    if (sheet.Cells[row, 1].Text == id.ToString())
                    {
                        sheet.DeleteRow(row);
                        break;
                    }
                }
                
                package.Save();
            }
            
            // Also delete associated BOM entries
            DeleteBOMByProductId(id);
        }

        #endregion

        #region Bill of Materials

        public List<BillOfMaterial> GetAllBOM()
        {
            lock (_lock)
            {
                var bomList = new List<BillOfMaterial>();
                using var package = new ExcelPackage(new FileInfo(_dataFilePath));
                var sheet = package.Workbook.Worksheets["BillOfMaterials"];
                
                if (sheet == null) return bomList;
                
                var rowCount = sheet.Dimension?.Rows ?? 0;
                for (int row = 2; row <= rowCount; row++)
                {
                    if (string.IsNullOrEmpty(sheet.Cells[row, 1].Text)) continue;
                    
                    bomList.Add(new BillOfMaterial
                    {
                        Id = Guid.TryParse(sheet.Cells[row, 1].Text, out var id) ? id : Guid.NewGuid(),
                        ProductId = Guid.TryParse(sheet.Cells[row, 2].Text, out var pid) ? pid : Guid.Empty,
                        MaterialId = Guid.TryParse(sheet.Cells[row, 3].Text, out var mid) ? mid : Guid.Empty,
                        Quantity = decimal.TryParse(sheet.Cells[row, 4].Text, out var qty) ? qty : 0
                    });
                }
                
                // Populate navigation properties
                var products = GetAllProducts();
                var materials = GetAllMaterials();
                
                foreach (var bom in bomList)
                {
                    var product = products.FirstOrDefault(p => p.Id == bom.ProductId);
                    var material = materials.FirstOrDefault(m => m.Id == bom.MaterialId);
                    
                    if (product != null)
                    {
                        bom.ProductName = product.Name;
                        bom.ProductCode = product.Code;
                    }
                    
                    if (material != null)
                    {
                        bom.MaterialName = material.Name;
                        bom.MaterialCode = material.Code;
                        bom.MaterialUnit = material.Unit;
                    }
                }
                
                return bomList;
            }
        }

        public List<BillOfMaterial> GetBOMByProductId(Guid productId)
        {
            return GetAllBOM().Where(b => b.ProductId == productId).ToList();
        }

        public void AddBOM(BillOfMaterial bom)
        {
            lock (_lock)
            {
                using var package = new ExcelPackage(new FileInfo(_dataFilePath));
                var sheet = package.Workbook.Worksheets["BillOfMaterials"];
                var rowCount = (sheet.Dimension?.Rows ?? 1) + 1;
                
                sheet.Cells[rowCount, 1].Value = bom.Id.ToString();
                sheet.Cells[rowCount, 2].Value = bom.ProductId.ToString();
                sheet.Cells[rowCount, 3].Value = bom.MaterialId.ToString();
                sheet.Cells[rowCount, 4].Value = bom.Quantity;
                
                package.Save();
            }
        }

        public void UpdateBOM(BillOfMaterial bom)
        {
            lock (_lock)
            {
                using var package = new ExcelPackage(new FileInfo(_dataFilePath));
                var sheet = package.Workbook.Worksheets["BillOfMaterials"];
                var rowCount = sheet.Dimension?.Rows ?? 0;
                
                for (int row = 2; row <= rowCount; row++)
                {
                    if (sheet.Cells[row, 1].Text == bom.Id.ToString())
                    {
                        sheet.Cells[row, 2].Value = bom.ProductId.ToString();
                        sheet.Cells[row, 3].Value = bom.MaterialId.ToString();
                        sheet.Cells[row, 4].Value = bom.Quantity;
                        break;
                    }
                }
                
                package.Save();
            }
        }

        public void DeleteBOM(Guid id)
        {
            lock (_lock)
            {
                using var package = new ExcelPackage(new FileInfo(_dataFilePath));
                var sheet = package.Workbook.Worksheets["BillOfMaterials"];
                var rowCount = sheet.Dimension?.Rows ?? 0;
                
                for (int row = 2; row <= rowCount; row++)
                {
                    if (sheet.Cells[row, 1].Text == id.ToString())
                    {
                        sheet.DeleteRow(row);
                        break;
                    }
                }
                
                package.Save();
            }
        }

        public void DeleteBOMByProductId(Guid productId)
        {
            lock (_lock)
            {
                using var package = new ExcelPackage(new FileInfo(_dataFilePath));
                var sheet = package.Workbook.Worksheets["BillOfMaterials"];
                var rowCount = sheet.Dimension?.Rows ?? 0;
                
                for (int row = rowCount; row >= 2; row--)
                {
                    if (sheet.Cells[row, 2].Text == productId.ToString())
                    {
                        sheet.DeleteRow(row);
                    }
                }
                
                package.Save();
            }
        }

        #endregion

        #region Stock Transactions

        public List<StockTransaction> GetAllTransactions()
        {
            lock (_lock)
            {
                var transactions = new List<StockTransaction>();
                using var package = new ExcelPackage(new FileInfo(_dataFilePath));
                var sheet = package.Workbook.Worksheets["StockTransactions"];
                
                if (sheet == null) return transactions;
                
                var rowCount = sheet.Dimension?.Rows ?? 0;
                for (int row = 2; row <= rowCount; row++)
                {
                    if (string.IsNullOrEmpty(sheet.Cells[row, 1].Text)) continue;
                    
                    transactions.Add(new StockTransaction
                    {
                        Id = Guid.TryParse(sheet.Cells[row, 1].Text, out var id) ? id : Guid.NewGuid(),
                        MaterialId = Guid.TryParse(sheet.Cells[row, 2].Text, out var mid) ? mid : Guid.Empty,
                        Type = Enum.TryParse<TransactionType>(sheet.Cells[row, 3].Text, out var type) ? type : TransactionType.IN,
                        Quantity = decimal.TryParse(sheet.Cells[row, 4].Text, out var qty) ? qty : 0,
                        Notes = sheet.Cells[row, 5].Text,
                        UserId = Guid.TryParse(sheet.Cells[row, 6].Text, out var uid) ? uid : Guid.Empty,
                        TransactionDate = DateTime.TryParse(sheet.Cells[row, 7].Text, out var date) ? date : DateTime.Now
                    });
                }
                
                // Populate navigation properties
                var materials = GetAllMaterials();
                var users = GetAllUsers();
                
                foreach (var trans in transactions)
                {
                    var material = materials.FirstOrDefault(m => m.Id == trans.MaterialId);
                    var user = users.FirstOrDefault(u => u.Id == trans.UserId);
                    
                    if (material != null)
                    {
                        trans.MaterialName = material.Name;
                        trans.MaterialCode = material.Code;
                    }
                    
                    if (user != null)
                    {
                        trans.UserName = user.FullName;
                    }
                }
                
                return transactions.OrderByDescending(t => t.TransactionDate).ToList();
            }
        }

        public List<StockTransaction> GetTransactionsByMaterialId(Guid materialId)
        {
            return GetAllTransactions().Where(t => t.MaterialId == materialId).ToList();
        }

        public void AddTransaction(StockTransaction transaction)
        {
            lock (_lock)
            {
                // First update the material stock
                var material = GetMaterialById(transaction.MaterialId);
                if (material != null)
                {
                    if (transaction.Type == TransactionType.IN)
                    {
                        material.CurrentStock += transaction.Quantity;
                    }
                    else
                    {
                        material.CurrentStock -= transaction.Quantity;
                        if (material.CurrentStock < 0) material.CurrentStock = 0;
                    }
                    UpdateMaterial(material);
                }
                
                // Then add the transaction record
                using var package = new ExcelPackage(new FileInfo(_dataFilePath));
                var sheet = package.Workbook.Worksheets["StockTransactions"];
                var rowCount = (sheet.Dimension?.Rows ?? 1) + 1;
                
                sheet.Cells[rowCount, 1].Value = transaction.Id.ToString();
                sheet.Cells[rowCount, 2].Value = transaction.MaterialId.ToString();
                sheet.Cells[rowCount, 3].Value = transaction.Type.ToString();
                sheet.Cells[rowCount, 4].Value = transaction.Quantity;
                sheet.Cells[rowCount, 5].Value = transaction.Notes;
                sheet.Cells[rowCount, 6].Value = transaction.UserId.ToString();
                sheet.Cells[rowCount, 7].Value = transaction.TransactionDate.ToString("O");
                
                package.Save();
            }
        }

        #endregion

        #region Stock Check

        public bool CanProduce(Guid productId, int quantity, out List<(Material material, decimal required, decimal available, decimal shortage)> shortages)
        {
            shortages = new List<(Material, decimal, decimal, decimal)>();
            
            var bomItems = GetBOMByProductId(productId);
            var materials = GetAllMaterials();
            
            bool canProduce = true;
            
            foreach (var bom in bomItems)
            {
                var material = materials.FirstOrDefault(m => m.Id == bom.MaterialId);
                if (material == null) continue;
                
                var required = bom.Quantity * quantity;
                if (material.CurrentStock < required)
                {
                    canProduce = false;
                    shortages.Add((material, required, material.CurrentStock, required - material.CurrentStock));
                }
            }
            
            return canProduce;
        }

        public int MaxProducible(Guid productId)
        {
            var bomItems = GetBOMByProductId(productId);
            if (!bomItems.Any()) return 0;
            
            var materials = GetAllMaterials();
            int maxQuantity = int.MaxValue;
            
            foreach (var bom in bomItems)
            {
                var material = materials.FirstOrDefault(m => m.Id == bom.MaterialId);
                if (material == null || bom.Quantity <= 0) continue;
                
                int possibleQuantity = (int)(material.CurrentStock / bom.Quantity);
                if (possibleQuantity < maxQuantity)
                {
                    maxQuantity = possibleQuantity;
                }
            }
            
            return maxQuantity == int.MaxValue ? 0 : maxQuantity;
        }

        #endregion

        #region Import/Export

        public void ExportToExcel(string filePath)
        {
            lock (_lock)
            {
                File.Copy(_dataFilePath, filePath, true);
            }
        }

        public void ImportFromExcel(string filePath)
        {
            // This is a simplified import that merges data
            // In production, you might want to add conflict resolution
            lock (_lock)
            {
                using var importPackage = new ExcelPackage(new FileInfo(filePath));
                
                // Import Materials
                var materialsSheet = importPackage.Workbook.Worksheets["Materials"];
                if (materialsSheet != null)
                {
                    var existingMaterials = GetAllMaterials().ToDictionary(m => m.Code);
                    var rowCount = materialsSheet.Dimension?.Rows ?? 0;
                    
                    for (int row = 2; row <= rowCount; row++)
                    {
                        if (string.IsNullOrEmpty(materialsSheet.Cells[row, 2].Text)) continue;
                        
                        var code = materialsSheet.Cells[row, 2].Text;
                        var material = new Material
                        {
                            Code = code,
                            Name = materialsSheet.Cells[row, 3].Text,
                            Description = materialsSheet.Cells[row, 4].Text,
                            Unit = materialsSheet.Cells[row, 5].Text,
                            CurrentStock = decimal.TryParse(materialsSheet.Cells[row, 6].Text, out var stock) ? stock : 0,
                            MinStockLevel = decimal.TryParse(materialsSheet.Cells[row, 7].Text, out var min) ? min : 0,
                            UnitPrice = decimal.TryParse(materialsSheet.Cells[row, 8].Text, out var price) ? price : 0,
                            Category = materialsSheet.Cells[row, 9].Text
                        };
                        
                        if (existingMaterials.ContainsKey(code))
                        {
                            material.Id = existingMaterials[code].Id;
                            UpdateMaterial(material);
                        }
                        else
                        {
                            AddMaterial(material);
                        }
                    }
                }
                
                // Import Products
                var productsSheet = importPackage.Workbook.Worksheets["Products"];
                if (productsSheet != null)
                {
                    var existingProducts = GetAllProducts().ToDictionary(p => p.Code);
                    var rowCount = productsSheet.Dimension?.Rows ?? 0;
                    
                    for (int row = 2; row <= rowCount; row++)
                    {
                        if (string.IsNullOrEmpty(productsSheet.Cells[row, 2].Text)) continue;
                        
                        var code = productsSheet.Cells[row, 2].Text;
                        var product = new Product
                        {
                            Code = code,
                            Name = productsSheet.Cells[row, 3].Text,
                            Description = productsSheet.Cells[row, 4].Text,
                            Category = productsSheet.Cells[row, 5].Text,
                            IsActive = bool.TryParse(productsSheet.Cells[row, 6].Text, out var active) && active
                        };
                        
                        if (existingProducts.ContainsKey(code))
                        {
                            product.Id = existingProducts[code].Id;
                            UpdateProduct(product);
                        }
                        else
                        {
                            AddProduct(product);
                        }
                    }
                }
            }
        }

        #endregion

        private bool ParseBool(ExcelRange cell)
        {
            var text = cell.Text.Trim();
            if (bool.TryParse(text, out var result)) return result;
            if (text == "1") return true;
            if (text == "0") return false;
            
            // Try GetValue
            try
            {
                return cell.GetValue<bool>();
            }
            catch
            {
                return false;
            }
        }
    }
}

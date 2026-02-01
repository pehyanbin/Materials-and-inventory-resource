using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using MIR.Models;
using MIR.Services;
using System.Collections.ObjectModel;

namespace MIR.ViewModels
{
    public partial class ProductsViewModel : ObservableObject
    {
        private readonly IExcelService _excelService;

        [ObservableProperty]
        private ObservableCollection<Product> _products = new();

        [ObservableProperty]
        private string _searchText = string.Empty;

        [ObservableProperty]
        private Product? _selectedProduct;

        [ObservableProperty]
        private bool _isAddingOrEditing;

        [ObservableProperty]
        private Product _editingProduct = new();

        // BOM related
        [ObservableProperty]
        private ObservableCollection<BillOfMaterial> _currentBOM = new();

        [ObservableProperty]
        private ObservableCollection<Material> _availableMaterials = new();

        [ObservableProperty]
        private Material? _selectedBomMaterial;

        [ObservableProperty]
        private decimal _bomQuantity = 1;

        public string[] Categories => ProductCategories.Categories;

        public ProductsViewModel(IExcelService excelService)
        {
            _excelService = excelService;
            LoadProducts();
            AvailableMaterials = new ObservableCollection<Material>(_excelService.GetAllMaterials());
        }

        [RelayCommand]
        private void LoadProducts()
        {
            var list = _excelService.GetAllProducts();
            if (!string.IsNullOrWhiteSpace(SearchText))
            {
                list = list.Where(p => 
                    p.Name.Contains(SearchText, StringComparison.OrdinalIgnoreCase) || 
                    p.Code.Contains(SearchText, StringComparison.OrdinalIgnoreCase)
                ).ToList();
            }
            Products = new ObservableCollection<Product>(list);
        }

        [RelayCommand]
        private void StartAddProduct()
        {
            EditingProduct = new Product { Code = $"PRD-{DateTime.Now:yyyyMMddHHmm}" };
            CurrentBOM = new ObservableCollection<BillOfMaterial>();
            IsAddingOrEditing = true;
        }

        [RelayCommand]
        private void StartEditProduct(Product product)
        {
            EditingProduct = new Product
            {
                Id = product.Id,
                Code = product.Code,
                Name = product.Name,
                Description = product.Description,
                Category = product.Category,
                IsActive = product.IsActive,
                CreatedAt = product.CreatedAt
            };
            
            CurrentBOM = new ObservableCollection<BillOfMaterial>(_excelService.GetBOMByProductId(product.Id));
            IsAddingOrEditing = true;
        }

        [RelayCommand]
        private void SaveProduct()
        {
            if (string.IsNullOrWhiteSpace(EditingProduct.Name) || string.IsNullOrWhiteSpace(EditingProduct.Code))
                return;

            if (_excelService.GetAllProducts().Any(p => p.Id == EditingProduct.Id))
            {
                _excelService.UpdateProduct(EditingProduct);
            }
            else
            {
                _excelService.AddProduct(EditingProduct);
            }

            // Save BOM
            _excelService.DeleteBOMByProductId(EditingProduct.Id);
            foreach (var bom in CurrentBOM)
            {
                bom.ProductId = EditingProduct.Id;
                _excelService.AddBOM(bom);
            }

            IsAddingOrEditing = false;
            LoadProducts();
        }

        [RelayCommand]
        private void AddToBOM()
        {
            if (SelectedBomMaterial == null || BomQuantity <= 0) return;

            // Check if already in BOM
            var existing = CurrentBOM.FirstOrDefault(b => b.MaterialId == SelectedBomMaterial.Id);
            if (existing != null)
            {
                existing.Quantity += BomQuantity;
            }
            else
            {
                CurrentBOM.Add(new BillOfMaterial(EditingProduct.Id, SelectedBomMaterial.Id, BomQuantity)
                {
                    MaterialName = SelectedBomMaterial.Name,
                    MaterialCode = SelectedBomMaterial.Code,
                    MaterialUnit = SelectedBomMaterial.Unit
                });
            }
            
            BomQuantity = 1;
        }

        [RelayCommand]
        private void RemoveFromBOM(BillOfMaterial bom)
        {
            CurrentBOM.Remove(bom);
        }

        [RelayCommand]
        private void CancelEdit()
        {
            IsAddingOrEditing = false;
        }

        [RelayCommand]
        private void DeleteProduct(Product product)
        {
            if (product == null) return;
            _excelService.DeleteProduct(product.Id);
            LoadProducts();
        }

        partial void OnSearchTextChanged(string value)
        {
            LoadProducts();
        }
    }
}

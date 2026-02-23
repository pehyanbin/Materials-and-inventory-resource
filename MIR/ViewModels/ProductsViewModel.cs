using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using MIR.Models;
using MIR.Services;
using System.Collections.ObjectModel;
using Microsoft.Win32;
using System.Windows;
using System.Threading.Tasks;

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

        [ObservableProperty]
        private bool _isImporting;

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

        [RelayCommand]
        private async Task ImportFromExcel()
        {
            if (IsImporting) return;

            var dialog = new OpenFileDialog
            {
                Title = "Import Products / BOM from Excel",
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                CheckFileExists = true
            };

            if (dialog.ShowDialog() != true) return;

            try
            {
                IsImporting = true;

                // Run import off the UI thread so the window stays responsive.
                await Task.Run(() => _excelService.ImportFromExcel(dialog.FileName));

                AvailableMaterials = new ObservableCollection<Material>(_excelService.GetAllMaterials());
                LoadProducts();

                MessageBox.Show(
                    "Excel import completed. Products, materials, and bill of materials were updated.",
                    "Import Complete",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Import failed: {ex.Message}",
                    "Import Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                IsImporting = false;
            }
        }

        partial void OnSearchTextChanged(string value)
        {
            LoadProducts();
        }
    }
}

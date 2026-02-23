using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using MIR.Models;
using MIR.Services;
using System.Collections.ObjectModel;

namespace MIR.ViewModels
{
    public partial class MaterialsViewModel : ObservableObject
    {
        private readonly IExcelService _excelService;

        [ObservableProperty]
        private ObservableCollection<Material> _materials = new();

        [ObservableProperty]
        private string _searchText = string.Empty;

        [ObservableProperty]
        private Material? _selectedMaterial;

        [ObservableProperty]
        private bool _isAddingOrEditing;

        [ObservableProperty]
        private Material _editingMaterial = new();

        public string[] Categories => MaterialCategories.Categories;

        [ObservableProperty]
        private string _currencySymbol = "$";

        public MaterialsViewModel(IExcelService excelService)
        {
            _excelService = excelService;
            CurrencySymbol = _excelService.GetSetting("Currency", "$");
            LoadMaterials();
        }

        [RelayCommand]
        private void LoadMaterials()
        {
            var list = _excelService.GetAllMaterials();
            if (!string.IsNullOrWhiteSpace(SearchText))
            {
                list = list.Where(m => 
                    m.Name.Contains(SearchText, StringComparison.OrdinalIgnoreCase) || 
                    m.Code.Contains(SearchText, StringComparison.OrdinalIgnoreCase)
                ).ToList();
            }
            Materials = new ObservableCollection<Material>(list);
        }

        [RelayCommand]
        private void StartAddMaterial()
        {
            EditingMaterial = new Material { Code = $"MAT-{DateTime.Now:yyyyMMddHHmm}" };
            IsAddingOrEditing = true;
        }

        [RelayCommand]
        private void StartEditMaterial(Material material)
        {
            // Clone the material for editing
            EditingMaterial = new Material
            {
                Id = material.Id,
                Code = material.Code,
                Name = material.Name,
                Description = material.Description,
                Unit = material.Unit,
                CurrentStock = material.CurrentStock,
                MinStockLevel = material.MinStockLevel,
                UnitPrice = material.UnitPrice,
                Category = material.Category,
                LastUpdated = material.LastUpdated
            };
            IsAddingOrEditing = true;
        }

        [RelayCommand]
        private void SaveMaterial()
        {
            if (string.IsNullOrWhiteSpace(EditingMaterial.Name) || string.IsNullOrWhiteSpace(EditingMaterial.Code))
                return;

            if (_excelService.GetAllMaterials().Any(m => m.Id == EditingMaterial.Id))
            {
                _excelService.UpdateMaterial(EditingMaterial);
            }
            else
            {
                _excelService.AddMaterial(EditingMaterial);
            }

            IsAddingOrEditing = false;
            LoadMaterials();
        }

        [RelayCommand]
        private void CancelEdit()
        {
            IsAddingOrEditing = false;
        }

        [RelayCommand]
        private void DeleteMaterial(Material material)
        {
            if (material == null) return;
            
            // In a real app, should show a confirmation dialog
            _excelService.DeleteMaterial(material.Id);
            LoadMaterials();
        }

        partial void OnSearchTextChanged(string value)
        {
            LoadMaterials();
        }
    }
}

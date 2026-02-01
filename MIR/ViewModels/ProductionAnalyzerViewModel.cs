using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using MIR.Models;
using MIR.Services;
using System.Collections.ObjectModel;

namespace MIR.ViewModels
{
    public partial class ProductionAnalyzerViewModel : ObservableObject
    {
        private readonly IExcelService _excelService;

        [ObservableProperty]
        private ObservableCollection<Product> _products = new();

        [ObservableProperty]
        private Product? _selectedProduct;

        [ObservableProperty]
        private int _targetQuantity = 10;

        [ObservableProperty]
        private ObservableCollection<AnalyzerEntry> _analysisResults = new();

        [ObservableProperty]
        private bool _hasShortages;

        [ObservableProperty]
        private string _totalShortageValue = "0.00";

        public ProductionAnalyzerViewModel(IExcelService excelService)
        {
            _excelService = excelService;
            LoadProducts();
        }

        private void LoadProducts()
        {
            var products = _excelService.GetAllProducts().Where(p => p.IsActive).ToList();
            Products = new ObservableCollection<Product>(products);
            if (Products.Any())
            {
                SelectedProduct = Products.First();
            }
        }

        [RelayCommand]
        private void Analyze()
        {
            if (SelectedProduct == null) return;

            var results = _excelService.GetProductionAnalysis(SelectedProduct.Id, TargetQuantity);
            
            AnalysisResults.Clear();
            decimal totalShortage = 0;
            bool missingAny = false;

            foreach (var item in results)
            {
                AnalysisResults.Add(new AnalyzerEntry
                {
                    MaterialName = item.Material.Name,
                    MaterialCode = item.Material.Code,
                    Unit = item.Material.Unit,
                    Required = item.Required,
                    Available = item.Available,
                    Shortage = item.Shortage,
                    ImpactPercentage = item.ImpactPercentage,
                    IsBottleneck = item.Shortage > 0
                });

                if (item.Shortage > 0)
                {
                    missingAny = true;
                    totalShortage += item.Shortage * item.Material.UnitPrice;
                }
            }

            HasShortages = missingAny;
            TotalShortageValue = totalShortage.ToString("N2");
        }

        partial void OnSelectedProductChanged(Product? value) => Analyze();
        partial void OnTargetQuantityChanged(int value) => Analyze();
    }

    public class AnalyzerEntry
    {
        public string MaterialName { get; set; } = string.Empty;
        public string MaterialCode { get; set; } = string.Empty;
        public string Unit { get; set; } = string.Empty;
        public decimal Required { get; set; }
        public decimal Available { get; set; }
        public decimal Shortage { get; set; }
        public decimal ImpactPercentage { get; set; }
        public bool IsBottleneck { get; set; }
    }
}

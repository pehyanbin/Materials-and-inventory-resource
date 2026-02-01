using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using MIR.Models;
using MIR.Services;
using System.Collections.ObjectModel;

namespace MIR.ViewModels
{
    public partial class ReportsViewModel : ObservableObject
    {
        private readonly IExcelService _excelService;

        [ObservableProperty]
        private string _selectedReportType = "Stock Report";

        [ObservableProperty]
        private ObservableCollection<string> _reportTypes = new() { "Stock Report", "Low Stock Report", "BOM Report" };

        [ObservableProperty]
        private ObservableCollection<Material> _materialsReport = new();

        [ObservableProperty]
        private ObservableCollection<Product> _products = new();

        [ObservableProperty]
        private Product? _selectedProduct;

        [ObservableProperty]
        private ObservableCollection<BillOfMaterial> _bomReport = new();

        [ObservableProperty]
        private bool _isProductSelectionVisible;

        [ObservableProperty]
        private decimal _totalInventoryValue;

        [ObservableProperty]
        private string _reportTitle = "Stock Report";

        [ObservableProperty]
        private string _currencySymbol = "$";

        public ReportsViewModel(IExcelService excelService)
        {
            _excelService = excelService;
            CurrencySymbol = _excelService.GetSetting("Currency", "$");
            LoadData();
        }

        private void LoadData()
        {
            GenerateReport();
        }

        [RelayCommand]
        private void GenerateReport()
        {
            ReportTitle = SelectedReportType;
            IsProductSelectionVisible = SelectedReportType == "BOM Report";

            switch (SelectedReportType)
            {
                case "Stock Report":
                    var materials = _excelService.GetAllMaterials();
                    MaterialsReport = new ObservableCollection<Material>(materials);
                    TotalInventoryValue = materials.Sum(m => m.CurrentStock * m.UnitPrice);
                    break;

                case "Low Stock Report":
                    var lowStock = _excelService.GetLowStockMaterials();
                    MaterialsReport = new ObservableCollection<Material>(lowStock);
                    TotalInventoryValue = lowStock.Sum(m => m.CurrentStock * m.UnitPrice);
                    break;

                case "BOM Report":
                    if (Products.Count == 0)
                    {
                        Products = new ObservableCollection<Product>(_excelService.GetAllProducts());
                        if (Products.Any()) SelectedProduct = Products.First();
                    }
                    UpdateBOMReport();
                    break;
            }
        }

        private void UpdateBOMReport()
        {
            if (SelectedProduct != null)
            {
                BomReport = new ObservableCollection<BillOfMaterial>(_excelService.GetBOMByProductId(SelectedProduct.Id));
            }
            else
            {
                BomReport.Clear();
            }
        }

        partial void OnSelectedReportTypeChanged(string value) => GenerateReport();
        partial void OnSelectedProductChanged(Product? value) => UpdateBOMReport();
    }
}

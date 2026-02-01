using CommunityToolkit.Mvvm.ComponentModel;
using MIR.Models;
using MIR.Services;
using System.Collections.ObjectModel;

namespace MIR.ViewModels
{
    public partial class DashboardViewModel : ObservableObject
    {
        private readonly IExcelService _excelService;

        [ObservableProperty]
        private int _totalMaterials;

        [ObservableProperty]
        private int _totalProducts;

        [ObservableProperty]
        private int _lowStockCount;

        [ObservableProperty]
        private ObservableCollection<Material> _lowStockMaterials = new();

        [ObservableProperty]
        private ObservableCollection<StockTransaction> _recentTransactions = new();

        public DashboardViewModel(IExcelService excelService)
        {
            _excelService = excelService;
            LoadStats();
        }

        private void LoadStats()
        {
            var materials = _excelService.GetAllMaterials();
            var products = _excelService.GetAllProducts();
            
            TotalMaterials = materials.Count;
            TotalProducts = products.Count;
            
            var lowStock = materials.Where(m => m.IsLowStock).ToList();
            LowStockCount = lowStock.Count;
            LowStockMaterials = new ObservableCollection<Material>(lowStock.Take(5));
            
            RecentTransactions = new ObservableCollection<StockTransaction>(_excelService.GetAllTransactions().Take(10));
        }
    }
}

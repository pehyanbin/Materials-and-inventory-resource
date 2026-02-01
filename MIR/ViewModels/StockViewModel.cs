using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using MIR.Models;
using MIR.Services;
using System.Collections.ObjectModel;

namespace MIR.ViewModels
{
    public partial class StockViewModel : ObservableObject
    {
        private readonly IExcelService _excelService;
        private readonly IAuthService _authService;

        [ObservableProperty]
        private ObservableCollection<Material> _materials = new();

        [ObservableProperty]
        private ObservableCollection<Product> _products = new();

        [ObservableProperty]
        private ObservableCollection<StockTransaction> _recentTransactions = new();

        [ObservableProperty]
        private Material? _selectedMaterial;

        [ObservableProperty]
        private decimal _transactionQuantity;

        [ObservableProperty]
        private string _transactionNotes = string.Empty;

        // Production Check
        [ObservableProperty]
        private Product? _checkProduct;

        [ObservableProperty]
        private int _checkQuantity = 1;

        [ObservableProperty]
        private string _checkResult = string.Empty;

        [ObservableProperty]
        private bool _canProduce;

        [ObservableProperty]
        private ObservableCollection<ShortageItem> _shortages = new();

        public StockViewModel(IExcelService excelService, IAuthService authService)
        {
            _excelService = excelService;
            _authService = authService;
            RefreshData();
        }

        private void RefreshData()
        {
            Materials = new ObservableCollection<Material>(_excelService.GetAllMaterials());
            Products = new ObservableCollection<Product>(_excelService.GetAllProducts().Where(p => p.IsActive));
            RecentTransactions = new ObservableCollection<StockTransaction>(_excelService.GetAllTransactions().Take(20));
        }

        [RelayCommand]
        private void AddStockIn()
        {
            ProcessTransaction(TransactionType.IN);
        }

        [RelayCommand]
        private void AddStockOut()
        {
            ProcessTransaction(TransactionType.OUT);
        }

        private void ProcessTransaction(TransactionType type)
        {
            if (SelectedMaterial == null || TransactionQuantity <= 0) return;

            var transaction = new StockTransaction(
                SelectedMaterial.Id,
                type,
                TransactionQuantity,
                _authService.CurrentUser?.Id ?? Guid.Empty,
                TransactionNotes
            );

            _excelService.AddTransaction(transaction);
            
            TransactionQuantity = 0;
            TransactionNotes = string.Empty;
            RefreshData();
        }

        [RelayCommand]
        private void CheckProduction()
        {
            if (CheckProduct == null) return;

            bool possible = _excelService.CanProduce(CheckProduct.Id, CheckQuantity, out var list);
            CanProduce = possible;
            
            var shortageList = new List<ShortageItem>();
            foreach (var item in list)
            {
                shortageList.Add(new ShortageItem
                {
                    MaterialName = item.material.Name,
                    Required = item.required,
                    Available = item.available,
                    Shortage = item.shortage,
                    Unit = item.material.Unit
                });
            }
            Shortages = new ObservableCollection<ShortageItem>(shortageList);

            if (possible)
            {
                CheckResult = $"Success! You have enough materials to produce {CheckQuantity} units of {CheckProduct.Name}.";
            }
            else
            {
                CheckResult = $"Shortage detected! You need more materials to produce {CheckQuantity} units of {CheckProduct.Name}.";
            }
        }

        [RelayCommand]
        private void CalculateMax()
        {
            if (CheckProduct == null) return;
            int max = _excelService.MaxProducible(CheckProduct.Id);
            CheckResult = $"Based on current stock, you can produce a maximum of {max} units of {CheckProduct.Name}.";
        }
    }

    public class ShortageItem
    {
        public string MaterialName { get; set; } = string.Empty;
        public decimal Required { get; set; }
        public decimal Available { get; set; }
        public decimal Shortage { get; set; }
        public string Unit { get; set; } = string.Empty;
    }
}

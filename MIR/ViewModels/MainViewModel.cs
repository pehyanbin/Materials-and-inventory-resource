using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using MIR.Services;
using System.Windows;

namespace MIR.ViewModels
{
    public partial class MainViewModel : ObservableObject
    {
        private readonly IAuthService _authService;
        private readonly IExcelService _excelService;

        [ObservableProperty]
        private string _title = "MIR - Materials & Inventory Resource";

        [ObservableProperty]
        private string _currentUserName = string.Empty;

        [ObservableProperty]
        private string _currentUserRole = string.Empty;

        [ObservableProperty]
        private object? _currentView;

        public MainViewModel(IAuthService authService, IExcelService excelService)
        {
            _authService = authService;
            _excelService = excelService;

            CurrentUserName = _authService.CurrentUser?.FullName ?? "Unknown User";
            CurrentUserRole = _authService.CurrentUser?.Role.ToString() ?? "Staff";

            // Default view
            Navigate("Dashboard");
        }

        [RelayCommand]
        private void Exit()
        {
            Application.Current.Shutdown();
        }

        [RelayCommand]
        private void Navigate(string target)
        {
            switch (target)
            {
                case "Dashboard":
                    CurrentView = new DashboardViewModel(_excelService);
                    break;
                case "Materials":
                    CurrentView = new MaterialsViewModel(_excelService);
                    break;
                case "Products":
                    CurrentView = new ProductsViewModel(_excelService);
                    break;
                case "Stock":
                    CurrentView = new StockViewModel(_excelService, _authService);
                    break;
                // Add other views here
                default:
                    CurrentView = null;
                    break;
            }
        }
    }
}

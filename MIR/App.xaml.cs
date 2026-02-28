using System.Windows;
using MIR.Services;
using MIR.ViewModels;
using MIR.Views;
using Microsoft.Extensions.DependencyInjection;

namespace MIR
{
    public partial class App : Application
    {
        private readonly ServiceProvider _serviceProvider;

        public App()
        {
            IServiceCollection services = new ServiceCollection();

            // Services
            services.AddSingleton<IExcelService, ExcelService>();
            services.AddSingleton<IAuthService, AuthService>();

            services.AddSingleton<MainViewModel>();

            // Windows
            services.AddSingleton<MainWindow>();

            _serviceProvider = services.BuildServiceProvider();
        }

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            // Initialize Excel storage
            var excelService = _serviceProvider.GetRequiredService<IExcelService>();
            excelService.Initialize();

            // Set default user and show main window
            var authService = _serviceProvider.GetRequiredService<IAuthService>();
            authService.SetDefaultUser();

            var mainWindow = _serviceProvider.GetRequiredService<MainWindow>();
            mainWindow.Show();
        }
    }
}

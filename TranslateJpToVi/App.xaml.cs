using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using TranslateJPToViLib;
using TranslateLib.Excel;
using TranslateLib.Interface;
using TranslateLib.PPT;

namespace TranslateJpToVi
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public static IHost? AppHost { get; private set; }
        public App()
        {
            AppHost = Host.CreateDefaultBuilder()
                .ConfigureServices((hostContext, services) =>
                {
                    services.AddSingleton<MainWindow>();
                    services.AddTransient<ITranslateExcel, TranslateExcelWithNpoi>();
                    services.AddTransient<ITranslate, TranslateWithGG>();
                    services.AddTransient<ITranslateFile, TranslatePPTWithSpire>();
                    var serviceProvider = services.BuildServiceProvider();
                    var logger = serviceProvider.GetService<ILogger<MainWindow>>();
                    services.AddSingleton(typeof(ILogger), logger);
                }).Build();
        }
        protected override async void OnStartup(StartupEventArgs e)
        {
            await AppHost!.StartAsync();

            var startupForm = AppHost.Services.GetRequiredService<MainWindow>();
            startupForm.Show();

            base.OnStartup(e);
        }
        protected override async void OnExit(ExitEventArgs e)
        {
            await AppHost!.StopAsync();
            base.OnExit(e);
        }
    }
}

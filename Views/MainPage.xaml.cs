using System.Windows;
using System.Windows.Controls;
using OofManager.Wpf.Models;
using OofManager.Wpf.ViewModels;

namespace OofManager.Wpf.Views;

public partial class MainPage : Page
{
    private readonly MainViewModel _vm;

    public MainPage(MainViewModel viewModel)
    {
        InitializeComponent();
        _vm = viewModel;
        DataContext = viewModel;
        Loaded += async (_, _) => await _vm.LoadAsync();
    }

    private void LoadTemplate_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn && btn.CommandParameter is OofTemplate template)
        {
            _vm.SelectedTemplate = template;
        }
    }
}

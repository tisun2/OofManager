using System.Windows.Controls;
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
}

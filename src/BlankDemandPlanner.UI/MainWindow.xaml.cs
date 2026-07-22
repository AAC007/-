using System.Windows;
using System.Windows.Controls;
using BlankDemandPlanner.UI.ViewModels;

namespace BlankDemandPlanner.UI;

public partial class MainWindow : Window
{
    public MainWindow(MainViewModel viewModel)
    {
        InitializeComponent();
        DataContext = viewModel;
    }

    private void DemandRow_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
    {
        if (sender is not DataGridRow { DataContext: DemandRow row } dataGridRow)
        {
            return;
        }

        if (FindParent<DataGrid>(dataGridRow)?.DataContext is DemandViewModel viewModel &&
            viewModel.ShowDetailsCommand.CanExecute(row))
        {
            viewModel.ShowDetailsCommand.Execute(row);
        }
    }

    private void DemandRowsGrid_SelectedCellsChanged(object sender, SelectedCellsChangedEventArgs e)
    {
        if (sender is DataGrid { CurrentItem: DemandRow row, DataContext: DemandViewModel viewModel } &&
            viewModel.ShowDetailsCommand.CanExecute(row))
        {
            viewModel.ShowDetailsCommand.Execute(row);
        }
    }

    private void LibraryRowsGrid_SelectedCellsChanged(object sender, SelectedCellsChangedEventArgs e)
    {
        if (sender is DataGrid { CurrentItem: LibraryRow row, DataContext: LibraryViewModel viewModel } &&
            viewModel.EditLibraryRowCommand.CanExecute(row))
        {
            viewModel.EditLibraryRowCommand.Execute(row);
        }
    }

    private void NsiRowsGrid_SelectedCellsChanged(object sender, SelectedCellsChangedEventArgs e)
    {
        if (sender is DataGrid { CurrentItem: NsiBlankRow row, DataContext: NormalizationViewModel viewModel } &&
            viewModel.LoadUsageCommand.CanExecute(row))
        {
            viewModel.LoadUsageCommand.Execute(row);
        }
    }

    private static T? FindParent<T>(DependencyObject child) where T : DependencyObject
    {
        var current = System.Windows.Media.VisualTreeHelper.GetParent(child);
        while (current is not null)
        {
            if (current is T typed)
            {
                return typed;
            }

            current = System.Windows.Media.VisualTreeHelper.GetParent(current);
        }

        return null;
    }
}

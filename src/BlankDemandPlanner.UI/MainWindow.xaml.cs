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

    private void LibraryRow_PreviewMouseRightButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (sender is not DataGridRow { DataContext: LibraryRow row } dataGridRow)
        {
            return;
        }

        dataGridRow.IsSelected = true;
        if (FindParent<DataGrid>(dataGridRow)?.DataContext is LibraryViewModel viewModel)
        {
            viewModel.SelectedRow = row;
        }
    }

    private void MskRow_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
    {
        if (sender is not DataGridRow { DataContext: MskLibraryRow row } dataGridRow)
        {
            return;
        }

        if (FindParent<DataGrid>(dataGridRow)?.DataContext is MskViewModel viewModel)
        {
            viewModel.SelectedRow = row;
        }
    }

    private void MskRowsGrid_SelectedCellsChanged(object sender, SelectedCellsChangedEventArgs e)
    {
        if (sender is DataGrid { CurrentItem: MskLibraryRow row, DataContext: MskViewModel viewModel })
        {
            viewModel.SelectedRow = row;
        }
    }

    private void MskRow_MouseLeftButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (sender is not DataGridRow { DataContext: MskLibraryRow row } dataGridRow)
        {
            return;
        }

        dataGridRow.IsSelected = true;
        if (FindParent<DataGrid>(dataGridRow)?.DataContext is MskViewModel viewModel &&
            viewModel.OpenDrawingCommand.CanExecute(row))
        {
            viewModel.SelectedRow = row;
            viewModel.OpenDrawingCommand.Execute(row);
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

    private void CalculationRow_PreviewMouseRightButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (sender is not DataGridRow { DataContext: CalculationMaterialRow row } dataGridRow)
        {
            return;
        }

        dataGridRow.IsSelected = true;
        if (FindParent<DataGrid>(dataGridRow)?.DataContext is CalculationViewModel viewModel)
        {
            viewModel.SelectedRow = row;
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

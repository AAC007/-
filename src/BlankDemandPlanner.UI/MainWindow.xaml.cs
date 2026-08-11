using System.ComponentModel;
using System.Reflection;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Forms.Integration;
using System.Windows.Input;
using BlankDemandPlanner.UI.ViewModels;
using PdfiumViewer;

namespace BlankDemandPlanner.UI;

public partial class MainWindow : Window
{
    private readonly MainViewModel viewModel;
    private bool isInitialLoadStarted;
    private WindowsFormsHost? mskPdfHost;
    private PdfViewer? mskPdfViewer;
    private PdfDocument? mskPdfDocument;
    private MskViewModel? subscribedMskViewModel;
    private WindowsFormsHost? blankSelectionPdfHost;
    private PdfViewer? blankSelectionPdfViewer;
    private PdfDocument? blankSelectionPdfDocument;
    private BlankSelectionViewModel? subscribedBlankSelectionViewModel;

    public MainWindow(MainViewModel viewModel)
    {
        InitializeComponent();
        this.viewModel = viewModel;
        DataContext = viewModel;
        SubscribeMskViewer(viewModel.Msk);
        SubscribeBlankSelectionViewer(viewModel.BlankSelection);
        Loaded += MainWindow_Loaded;
        Closed += (_, _) =>
        {
            UnsubscribeMskViewer();
            DisposeMskPdfViewer();
            UnsubscribeBlankSelectionViewer();
            DisposeBlankSelectionPdfViewer();
        };
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

    private void DemandRow_PreviewMouseRightButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (sender is not DataGridRow { DataContext: DemandRow row } dataGridRow)
        {
            return;
        }

        dataGridRow.IsSelected = true;
        if (FindParent<DataGrid>(dataGridRow)?.DataContext is DemandViewModel viewModel)
        {
            viewModel.SelectedRow = row;
        }
    }

    private void LibraryRowsGrid_SelectedCellsChanged(object sender, SelectedCellsChangedEventArgs e)
    {
        if (sender is DataGrid { CurrentItem: LibraryRow row, DataContext: LibraryViewModel viewModel })
        {
            viewModel.SelectedRow = row;
        }
    }

    private void LibraryRow_PreviewMouseRightButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (sender is not DataGridRow { DataContext: LibraryRow row } dataGridRow)
        {
            return;
        }

        if (FindParent<DataGrid>(dataGridRow)?.DataContext is LibraryViewModel viewModel)
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

        if (FindParent<DataGrid>(dataGridRow)?.DataContext is MskViewModel viewModel)
        {
            viewModel.SelectedRow = row;
            if (viewModel.OpenDrawingCommand.CanExecute(row))
            {
                viewModel.OpenDrawingCommand.Execute(row);
            }
        }
    }

    private void NavigationItem_PreviewMouseRightButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        e.Handled = true;
        NavigationItem_MouseRightButtonUp(sender, e);
    }

    private void NavigationItem_MouseRightButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (sender is not ListBoxItem { DataContext: NavigationItem navigationItem })
        {
            return;
        }

        var menuItem = new MenuItem { Header = "Открыть в новом окне" };
        menuItem.Click += (_, _) => OpenNavigationItemInNewWindow(navigationItem);
        var menu = new ContextMenu
        {
            PlacementTarget = navigationItem is null ? null : sender as UIElement
        };
        menu.Items.Add(menuItem);
        menu.IsOpen = true;
        e.Handled = true;
    }

    private void OpenNavigationItemInNewWindow(NavigationItem navigationItem)
    {
        if (navigationItem.Page is null)
        {
            return;
        }

        var template = TryFindResource(new DataTemplateKey(navigationItem.Page.GetType())) as DataTemplate;
        var content = new ContentControl
        {
            Content = navigationItem.Page,
            ContentTemplate = template
        };

        var window = new Window
        {
            Title = $"Планирование ЦМО - {navigationItem.Title}",
            Width = Math.Max(1000, ActualWidth * 0.82),
            Height = Math.Max(680, ActualHeight * 0.82),
            MinWidth = 900,
            MinHeight = 620,
            WindowStartupLocation = WindowStartupLocation.CenterScreen,
            ShowInTaskbar = true,
            DataContext = viewModel,
            Content = content,
            FontFamily = FontFamily,
            FontSize = FontSize,
            Icon = Icon
        };

        window.Show();
    }

    private void EditableComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (sender is not System.Windows.Controls.ComboBox { IsEditable: true, SelectedItem: not null } comboBox)
        {
            return;
        }

        var text = comboBox.SelectedItem is string stringValue
            ? stringValue
            : ReadDisplayText(comboBox.SelectedItem, comboBox.DisplayMemberPath);
        if (string.IsNullOrWhiteSpace(text))
        {
            return;
        }

        comboBox.SetCurrentValue(System.Windows.Controls.ComboBox.TextProperty, text);
        comboBox.GetBindingExpression(System.Windows.Controls.ComboBox.TextProperty)?.UpdateSource();
        comboBox.Dispatcher.BeginInvoke(new Action(() =>
        {
            comboBox.SetCurrentValue(System.Windows.Controls.ComboBox.TextProperty, text);
            comboBox.GetBindingExpression(System.Windows.Controls.ComboBox.TextProperty)?.UpdateSource();
            comboBox.IsDropDownOpen = false;
        }), System.Windows.Threading.DispatcherPriority.Background);
    }

    private static string ReadDisplayText(object value, string? displayMemberPath)
    {
        if (!string.IsNullOrWhiteSpace(displayMemberPath))
        {
            var property = value.GetType().GetProperty(displayMemberPath);
            var propertyValue = property?.GetValue(value)?.ToString();
            if (!string.IsNullOrWhiteSpace(propertyValue))
            {
                return propertyValue;
            }
        }

        return value.ToString() ?? string.Empty;
    }

    private void CopySelectedCells_CanExecute(object sender, CanExecuteRoutedEventArgs e)
    {
        if (Keyboard.FocusedElement is System.Windows.Controls.Primitives.TextBoxBase)
        {
            e.CanExecute = false;
            return;
        }

        var grid = ResolveFocusedDataGrid();
        e.CanExecute = grid?.SelectedCells.Count > 0;
        e.Handled = e.CanExecute;
    }

    private void CopySelectedCells_Executed(object sender, ExecutedRoutedEventArgs e)
    {
        var grid = ResolveFocusedDataGrid();
        if (grid is null || grid.SelectedCells.Count == 0)
        {
            return;
        }

        var text = BuildSelectedCellsClipboardText(grid);
        if (!string.IsNullOrEmpty(text))
        {
            System.Windows.Clipboard.SetText(text);
        }

        e.Handled = true;
    }

    private DataGrid? ResolveFocusedDataGrid()
    {
        if (Keyboard.FocusedElement is not DependencyObject focused)
        {
            return FindSelectedDataGrid(this);
        }

        return focused is DataGrid dataGrid
            ? dataGrid
            : FindParent<DataGrid>(focused) ?? FindSelectedDataGrid(this);
    }

    private static DataGrid? FindSelectedDataGrid(DependencyObject root)
    {
        if (root is DataGrid { SelectedCells.Count: > 0 } grid)
        {
            return grid;
        }

        var count = System.Windows.Media.VisualTreeHelper.GetChildrenCount(root);
        for (var i = 0; i < count; i++)
        {
            var child = System.Windows.Media.VisualTreeHelper.GetChild(root, i);
            var result = FindSelectedDataGrid(child);
            if (result is not null)
            {
                return result;
            }
        }

        return null;
    }

    private static string BuildSelectedCellsClipboardText(DataGrid grid)
    {
        var cells = grid.SelectedCells
            .Where(x => x.Item is not null && x.Column is not null)
            .ToList();
        if (cells.Count == 0)
        {
            return string.Empty;
        }

        var itemOrder = grid.Items.Cast<object>()
            .Select((item, index) => (item, index))
            .ToDictionary(x => x.item, x => x.index);
        var rows = cells
            .Select(x => x.Item)
            .Distinct()
            .OrderBy(x => itemOrder.GetValueOrDefault(x, int.MaxValue))
            .ToList();
        var columns = cells
            .Select(x => x.Column)
            .Distinct()
            .OrderBy(x => x.DisplayIndex)
            .ToList();
        var selected = cells.ToLookup(x => (x.Item, x.Column));
        var builder = new StringBuilder();
        foreach (var row in rows)
        {
            if (builder.Length > 0)
            {
                builder.AppendLine();
            }

            for (var i = 0; i < columns.Count; i++)
            {
                if (i > 0)
                {
                    builder.Append('\t');
                }

                var column = columns[i];
                if (selected.Contains((row, column)))
                {
                    builder.Append(GetCellClipboardValue(column, row));
                }
            }
        }

        return builder.ToString();
    }

    private static string GetCellClipboardValue(DataGridColumn column, object item)
    {
        if (column.ClipboardContentBinding is System.Windows.Data.Binding clipboardBinding)
        {
            return FormatClipboardValue(ReadPropertyPath(item, clipboardBinding.Path?.Path), clipboardBinding.StringFormat);
        }

        if (column is DataGridBoundColumn { Binding: System.Windows.Data.Binding binding })
        {
            return FormatClipboardValue(ReadPropertyPath(item, binding.Path?.Path), binding.StringFormat);
        }

        var content = column.GetCellContent(item);
        return content switch
        {
            TextBlock textBlock => textBlock.Text,
            System.Windows.Controls.TextBox textBox => textBox.Text,
            ContentControl contentControl => contentControl.Content?.ToString() ?? string.Empty,
            _ => content?.ToString() ?? string.Empty
        };
    }

    private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        if (isInitialLoadStarted)
        {
            return;
        }

        isInitialLoadStarted = true;
        try
        {
            await viewModel.StartupLoadAsync();
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show(
                $"Не удалось загрузить данные приложения.\n{ex.GetBaseException().Message}",
                "Планирование ЦМО",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private static object? ReadPropertyPath(object item, string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return item;
        }

        object? current = item;
        foreach (var part in path.Split('.'))
        {
            if (current is null)
            {
                return null;
            }

            var property = current.GetType().GetProperty(part, BindingFlags.Instance | BindingFlags.Public);
            current = property?.GetValue(current);
        }

        return current;
    }

    private static string FormatClipboardValue(object? value, string? stringFormat)
    {
        if (value is null)
        {
            return string.Empty;
        }

        if (!string.IsNullOrWhiteSpace(stringFormat))
        {
            return string.Format(System.Globalization.CultureInfo.CurrentCulture, stringFormat, value);
        }

        return value.ToString() ?? string.Empty;
    }

    private void SubscribeMskViewer(MskViewModel viewModel)
    {
        subscribedMskViewModel = viewModel;
        subscribedMskViewModel.PropertyChanged += MskViewModel_PropertyChanged;
    }

    private void UnsubscribeMskViewer()
    {
        if (subscribedMskViewModel is not null)
        {
            subscribedMskViewModel.PropertyChanged -= MskViewModel_PropertyChanged;
            subscribedMskViewModel = null;
        }
    }

    private void MskPdfHost_Loaded(object sender, RoutedEventArgs e)
    {
        if (sender is not WindowsFormsHost host)
        {
            return;
        }

        try
        {
            mskPdfHost = host;
            mskPdfViewer = new PdfViewer { Dock = System.Windows.Forms.DockStyle.Fill };
            mskPdfHost.Child = mskPdfViewer;
        }
        catch (Exception ex)
        {
            mskPdfViewer?.Dispose();
            mskPdfViewer = null;
            mskPdfHost = null;
            if (subscribedMskViewModel is not null)
            {
                subscribedMskViewModel.DrawingStatusText = $"PDF-viewer не запущен: {ex.GetBaseException().Message}";
            }
            return;
        }

        if (subscribedMskViewModel is not null)
        {
            ShowMskPdf(subscribedMskViewModel.DrawingViewerSource);
        }
    }

    private void MskPdfHost_Unloaded(object sender, RoutedEventArgs e)
    {
        if (ReferenceEquals(sender, mskPdfHost))
        {
            DisposeMskPdfViewer();
        }
    }

    private void MskViewModel_PropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(MskViewModel.DrawingViewerSource) && sender is MskViewModel viewModel)
        {
            ShowMskPdf(viewModel.DrawingViewerSource);
        }
    }

    private void ShowMskPdf(Uri? source)
    {
        if (mskPdfViewer is null)
        {
            return;
        }

        mskPdfViewer.Document = null;
        mskPdfDocument?.Dispose();
        mskPdfDocument = null;

        if (source is null || !source.IsFile || !System.IO.File.Exists(source.LocalPath))
        {
            return;
        }

        try
        {
            mskPdfDocument = PdfDocument.Load(source.LocalPath);
            mskPdfViewer.Document = mskPdfDocument;
            mskPdfViewer.ZoomMode = PdfViewerZoomMode.FitWidth;
        }
        catch (Exception ex)
        {
            mskPdfDocument?.Dispose();
            mskPdfDocument = null;
            if (subscribedMskViewModel is not null)
            {
                subscribedMskViewModel.DrawingStatusText = $"PDF найден, но viewer не смог открыть файл: {ex.GetBaseException().Message}";
            }
        }
    }

    private void MskOpenPdfFullWindow_Click(object sender, RoutedEventArgs e)
    {
        var source = subscribedMskViewModel?.DrawingViewerSource;
        if (source is null || !source.IsFile || !System.IO.File.Exists(source.LocalPath))
        {
            if (subscribedMskViewModel is not null)
            {
                subscribedMskViewModel.DrawingStatusText = "PDF-чертеж не выбран.";
            }

            return;
        }

        PdfViewer? fullViewer = null;
        PdfDocument? fullDocument = null;
        WindowsFormsHost? fullHost = null;
        try
        {
            fullDocument = PdfDocument.Load(source.LocalPath);
            fullViewer = new PdfViewer
            {
                Dock = System.Windows.Forms.DockStyle.Fill,
                Document = fullDocument,
                ZoomMode = PdfViewerZoomMode.FitWidth
            };
            fullHost = new WindowsFormsHost { Child = fullViewer };

            var window = new Window
            {
                Title = $"PDF-чертеж {System.IO.Path.GetFileName(source.LocalPath)}",
                Owner = this,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                WindowState = WindowState.Maximized,
                Content = fullHost
            };
            window.Closed += (_, _) =>
            {
                fullViewer.Document = null;
                fullHost.Child = null;
                fullViewer.Dispose();
                fullDocument.Dispose();
            };
            window.Show();
        }
        catch (Exception ex)
        {
            fullViewer?.Dispose();
            fullDocument?.Dispose();
            if (subscribedMskViewModel is not null)
            {
                subscribedMskViewModel.DrawingStatusText = $"PDF найден, но полноэкранный viewer не смог открыть файл: {ex.GetBaseException().Message}";
            }
        }
    }

    private void DisposeMskPdfViewer()
    {
        if (mskPdfViewer is not null)
        {
            mskPdfViewer.Document = null;
            mskPdfViewer.Dispose();
            mskPdfViewer = null;
        }

        mskPdfDocument?.Dispose();
        mskPdfDocument = null;
        if (mskPdfHost is not null)
        {
            mskPdfHost.Child = null;
            mskPdfHost = null;
        }
    }

    private void SubscribeBlankSelectionViewer(BlankSelectionViewModel viewModel)
    {
        subscribedBlankSelectionViewModel = viewModel;
        subscribedBlankSelectionViewModel.PropertyChanged += BlankSelectionViewModel_PropertyChanged;
    }

    private void UnsubscribeBlankSelectionViewer()
    {
        if (subscribedBlankSelectionViewModel is not null)
        {
            subscribedBlankSelectionViewModel.PropertyChanged -= BlankSelectionViewModel_PropertyChanged;
            subscribedBlankSelectionViewModel = null;
        }
    }

    private void BlankSelectionPdfHost_Loaded(object sender, RoutedEventArgs e)
    {
        if (sender is not WindowsFormsHost host)
        {
            return;
        }

        try
        {
            blankSelectionPdfHost = host;
            blankSelectionPdfViewer = new PdfViewer { Dock = System.Windows.Forms.DockStyle.Fill };
            blankSelectionPdfHost.Child = blankSelectionPdfViewer;
        }
        catch (Exception ex)
        {
            blankSelectionPdfViewer?.Dispose();
            blankSelectionPdfViewer = null;
            blankSelectionPdfHost = null;
            if (subscribedBlankSelectionViewModel is not null)
            {
                subscribedBlankSelectionViewModel.DrawingStatusText = $"PDF-viewer не запущен: {ex.GetBaseException().Message}";
            }
            return;
        }

        if (subscribedBlankSelectionViewModel is not null)
        {
            ShowBlankSelectionPdf(subscribedBlankSelectionViewModel.DrawingViewerSource);
        }
    }

    private void BlankSelectionPdfHost_Unloaded(object sender, RoutedEventArgs e)
    {
        if (ReferenceEquals(sender, blankSelectionPdfHost))
        {
            DisposeBlankSelectionPdfViewer();
        }
    }

    private void BlankSelectionViewModel_PropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(BlankSelectionViewModel.DrawingViewerSource) && sender is BlankSelectionViewModel viewModel)
        {
            ShowBlankSelectionPdf(viewModel.DrawingViewerSource);
        }
    }

    private void ShowBlankSelectionPdf(Uri? source)
    {
        if (blankSelectionPdfViewer is null)
        {
            return;
        }

        blankSelectionPdfViewer.Document = null;
        blankSelectionPdfDocument?.Dispose();
        blankSelectionPdfDocument = null;

        if (source is null || !source.IsFile || !System.IO.File.Exists(source.LocalPath))
        {
            return;
        }

        try
        {
            blankSelectionPdfDocument = PdfDocument.Load(source.LocalPath);
            blankSelectionPdfViewer.Document = blankSelectionPdfDocument;
            blankSelectionPdfViewer.ZoomMode = PdfViewerZoomMode.FitWidth;
        }
        catch (Exception ex)
        {
            blankSelectionPdfDocument?.Dispose();
            blankSelectionPdfDocument = null;
            if (subscribedBlankSelectionViewModel is not null)
            {
                subscribedBlankSelectionViewModel.DrawingStatusText = $"PDF найден, но viewer не смог открыть файл: {ex.GetBaseException().Message}";
            }
        }
    }

    private void DisposeBlankSelectionPdfViewer()
    {
        if (blankSelectionPdfViewer is not null)
        {
            blankSelectionPdfViewer.Document = null;
            blankSelectionPdfViewer.Dispose();
            blankSelectionPdfViewer = null;
        }

        blankSelectionPdfDocument?.Dispose();
        blankSelectionPdfDocument = null;
        if (blankSelectionPdfHost is not null)
        {
            blankSelectionPdfHost.Child = null;
            blankSelectionPdfHost = null;
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

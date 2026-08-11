using Microsoft.Win32;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;

namespace BlankDemandPlanner.UI.Services;

public interface IFileDialogService
{
    string? OpenExcelFile();
    string? SelectFolder();
}

public sealed class FileDialogService : IFileDialogService
{
    public string? OpenExcelFile()
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|Все файлы (*.*)|*.*",
            CheckFileExists = true
        };
        return dialog.ShowDialog() == true ? dialog.FileName : null;
    }

    public string? SelectFolder()
    {
        var dialog = new OpenFolderDialog();
        return dialog.ShowDialog() == true ? dialog.FolderName : null;
    }
}

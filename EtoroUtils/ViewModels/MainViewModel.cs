using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Platform.Storage;
using CommunityToolkit.Mvvm.ComponentModel;
using System.Diagnostics;
using System.Threading.Tasks;

namespace EtoroUtils.ViewModels;

public partial class MainViewModel : ViewModelBase
{
    public static FilePickerFileType FileTypeXLSX { get; } = new("Excel(xlsx)") {
        Patterns = ["*.xlsx"],
        //AppleUniformTypeIdentifiers = new[] { "public.image" },
        //MimeTypes = new[] { "image/*" }
    };

    [ObservableProperty]
    private string status = "";

    [ObservableProperty]
    private bool working = false;

    public async void Process() {
        
        //I don't care
        var lt = Avalonia.Application.Current.ApplicationLifetime as ClassicDesktopStyleApplicationLifetime;
        var sp = lt.MainWindow.StorageProvider;
        var folder = await sp.TryGetWellKnownFolderAsync(WellKnownFolder.Desktop);
        //This can also be applied for SaveFilePicker.
        var files = await sp.OpenFilePickerAsync(new FilePickerOpenOptions() {
            AllowMultiple = false,
            Title = "title",
            //You can add either custom or from the built-in file types. See "Defining custom file types" on how to create a custom one.
            FileTypeFilter = [FileTypeXLSX],
            SuggestedStartLocation = folder,
        });

        var path = files[0].Path.LocalPath;
        Status = $"Processing: {path}";
        EtoroStatementProcessor etoroStatementProcessor = new();
        Working = true;
        await Task.Run(() => etoroStatementProcessor.Process(path));
        Working = false;
        Status = $"Done: {path}";



        // combine the arguments together
        // it doesn't matter if there is a space after ','
        string argument = $"/select, \"{path}\" ";

        System.Diagnostics.Process.Start("explorer.exe", argument);

    }
}

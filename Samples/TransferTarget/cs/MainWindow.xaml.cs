using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using Microsoft.UI.Dispatching;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Media.Imaging;
using Windows.ApplicationModel.DataTransfer;
using Windows.Foundation.Metadata;
using Windows.Storage;
using Windows.Storage.Pickers;
using WinRT.Interop;

namespace Sample
{
    /// <summary>
    /// Represents a discovered transfer target for data binding.
    /// </summary>
    public sealed class TransferTargetItem : INotifyPropertyChanged
    {
        private readonly TransferTargetWatcher _watcher;
        private TransferTarget _target;
        private readonly string _id;
        private BitmapImage? _logo;
        private bool _forceFallback;

        public event PropertyChangedEventHandler? PropertyChanged;

        public TransferTargetItem(TransferTargetWatcher watcher, TransferTarget target)
        {
            _watcher = watcher;
            _target = target;
            _id = target.Id;
        }

        public string Label => _target.Label;

        public ImageSource? Logo
        {
            get
            {
                if (_forceFallback)
                {
                    return null;
                }

                if (_logo == null)
                {
                    StartLoadLogo();
                }
                return _logo;
            }
        }

        public bool IsEnabled => _target.IsEnabled;

        public string Id => _id;

        private void ReportLogoFailed()
        {
            // When the bitmap logo file fails to load, remember that it failed
            // and raise the logo changed events to update the UI with the fallback logo.
            _forceFallback = true;

            // Defer raising the event because we don't want to raise it while in the middle of a bind operation.
            DispatcherQueue.GetForCurrentThread().TryEnqueue(() =>
            {
                RaisePropertyChanged(nameof(Logo));
            });
        }

        private async void StartLoadLogo()
        {
            try
            {
                _logo = new BitmapImage();

                _logo.ImageFailed += (s, e) =>
                {
                    ReportLogoFailed();
                };

                var icon = _target.DisplayIcon;
                if (icon == null)
                {
                    ReportLogoFailed();
                    return;
                }

                var stream = await icon.OpenReadAsync();
                _logo.SetSource(stream);
            }
            catch
            {
                ReportLogoFailed();
            }
        }

        private static Windows.UI.WindowId WindowIdFromElement(UIElement element) =>
            new Windows.UI.WindowId { Value = element.XamlRoot.ContentIslandEnvironment.AppWindowId.Value };

        public async void Invoke(object sender, RoutedEventArgs e)
        {
            var result = await _watcher.TransferToAsync(_target, WindowIdFromElement((UIElement)sender));
        }

        public void Update(TransferTarget target)
        {
            _forceFallback = false;
            _target = target;
            RaisePropertyChanged(nameof(Label));
            RaisePropertyChanged(nameof(Logo));
            RaisePropertyChanged(nameof(IsEnabled));
        }

        private void RaisePropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    /// <summary>
    /// Main window that demonstrates the TransferTarget API.
    /// </summary>
    public sealed partial class MainWindow : Window
    {
        // Store selected files for later use with dataPackage
        private IReadOnlyList<IStorageItem> _selectedFiles = new List<IStorageItem>();

        // Observable collection of found targets for data binding to the ListView
        public ObservableCollection<TransferTargetItem> DiscoveredTargets { get; } = new ObservableCollection<TransferTargetItem>();

        // TransferTargetWatcher members
        private TransferTargetWatcher? _watcher;
        private bool _watcherRunning;

        public MainWindow()
        {
            InitializeComponent();

            // Disable the feature if the TransferTargetWatcher is not available.
            if (!ApiInformation.IsTypePresent("Windows.ApplicationModel.DataTransfer.TransferTargetWatcher"))
            {
                NotSupportedPanel.Visibility = Visibility.Visible;
                ConfigurationPanel.Visibility = Visibility.Collapsed;
            }
        }

        // Helper methods for binding
        public static double EnabledOpacity(bool isEnabled) => isEnabled ? 1.0 : 0.5;
        public static bool IsNull(object? value) => value == null;
        public static bool IsNotNull(object? value) => value != null;

        private void OnTargetAdded(TransferTargetWatcher sender, TransferTargetChangedEventArgs args)
        {
            var target = args.Target;

            // Switch to UI thread
            DispatcherQueue.TryEnqueue(() =>
            {
                // Check that the event came from the watcher we are using.
                // This avoids a race condition where the event is in flight when the watcher is stopped.
                if (sender == _watcher)
                {
                    DiscoveredTargets.Add(new TransferTargetItem(sender, target));
                }
            });
        }

        private TransferTargetItem? FindTargetItemById(string id)
        {
            return DiscoveredTargets.FirstOrDefault(item => item.Id == id);
        }

        private void OnTargetRemoved(TransferTargetWatcher sender, TransferTargetChangedEventArgs args)
        {
            var target = args.Target;

            // Switch to UI thread
            DispatcherQueue.TryEnqueue(() =>
            {
                // Check that the event came from the watcher we are using.
                // This avoids a race condition where the event is in flight when the watcher is stopped.
                if (sender == _watcher)
                {
                    // If the target is in the list, remove it.
                    // Note that the "target" is not necessarily the same object as the one passed to the Added event.
                    // You have to match it up by the ID.
                    var item = FindTargetItemById(target.Id);
                    if (item != null)
                    {
                        DiscoveredTargets.Remove(item);
                    }
                }
            });
        }

        private void OnTargetUpdated(TransferTargetWatcher sender, TransferTargetChangedEventArgs args)
        {
            var target = args.Target;

            // Switch to UI thread
            DispatcherQueue.TryEnqueue(() =>
            {
                // Check that the event came from the watcher we are using.
                // This avoids a race condition where the event is in flight when the watcher is stopped.
                if (sender == _watcher)
                {
                    // If the target is in the list, update it.
                    // Note that the "target" is not necessarily the same object as the one passed to the Added event.
                    // You have to match it up by the ID.
                    var item = FindTargetItemById(target.Id);
                    if (item != null)
                    {
                        item.Update(target);
                    }
                }
            });
        }

        private async void SelectFilesButton_Click(object sender, RoutedEventArgs e)
        {
            // Create and configure a file picker.
            var picker = new FileOpenPicker();
            picker.ViewMode = PickerViewMode.List;
            picker.SuggestedStartLocation = PickerLocationId.DocumentsLibrary;
            picker.FileTypeFilter.Add("*");

            // Initialize the picker with our window handle.
            InitializeWithWindow.Initialize(picker, WindowNative.GetWindowHandle(this));

            // Show the picker and remember the selected files.
            var files = await picker.PickMultipleFilesAsync();

            _selectedFiles = files;

            // Update the label that summarizes how many files were selected.
            if (_selectedFiles.Count > 0)
            {
                FilesCountLabel.Text = $"{_selectedFiles.Count} file(s) selected";
            }
            else
            {
                FilesCountLabel.Text = "No files selected";
            }
        }

        private void DiscoverTargetsButton_Click(object sender, RoutedEventArgs e)
        {
            // Create a new data package
            var dataPackage = new DataPackage();

            // Fill the DataPackage with requested data.

            // StorageItems
            if (_selectedFiles.Count > 0)
            {
                dataPackage.SetStorageItems(_selectedFiles);
            }

            // Link
            if (LinkComboBox.SelectedIndex > 0)
            {
                var uri = new Uri((string)LinkComboBox.SelectedItem);
                dataPackage.SetApplicationLink(uri);
                dataPackage.SetWebLink(uri);
            }

            // Text
            if (TextCheckBox.IsChecked == true)
            {
                dataPackage.SetText(TextInput.Text);
            }

            // HTML
            if (HtmlCheckBox.IsChecked == true)
            {
                dataPackage.SetHtmlFormat(HtmlFormatHelper.CreateHtmlFormat(HtmlInput.Text));
            }

            // Custom Data Format
            if (CustomDataFormatCheckBox.IsChecked == true)
            {
                dataPackage.SetData(CustomDataFormatNameInput.Text, CustomDataContentInput.Text);
            }

            // Set additional data package properties for illustration.
            // For demonstration purposes, these are hard-coded.
            var properties = dataPackage.Properties;
            properties.Title = "Sample Transfer Target Test";
            properties.Description = "Testing TransferTargetWatcher API with multiple data formats";

            properties.ContentSourceWebLink = new Uri("https://www.microsoft.com");
            properties.ContentSourceApplicationLink = new Uri("sampleapp://test");
            properties.ApplicationListingUri = new Uri("https://www.microsoft.com/store");

            // Add custom properties for additional metadata.
            properties["Author"] = "Test User";
            properties["Version"] = "1.0.0";
            properties["Category"] = "Testing";

            // Create TransferTargetDiscoveryOptions with the data package view
            var discoveryOptions = new TransferTargetDiscoveryOptions(dataPackage.GetView());

            // Maximum number of app targets.
            discoveryOptions.MaxAppTargets = (int)NumberOfTargetsComboBox.SelectedItem;

            // Set AllowedTargetAppIds from the app filter input (comma-separated list) only if checkbox is checked
            if (AppFilterCheckBox.IsChecked!.Value)
            {
                // Parse comma-separated AppIds into an array
                var appIds = AppFilterInput.Text
                    .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

                discoveryOptions.AllowedTargetAppIds = appIds;
            }

            // Create the watcher
            _watcher = TransferTarget.CreateWatcher(discoveryOptions);

            // Register for watcher events
            _watcher.Added += OnTargetAdded;
            _watcher.Removed += OnTargetRemoved;
            _watcher.Updated += OnTargetUpdated;

            // Start the watcher
            _watcher.Start();
            _watcherRunning = true;

            // Go to Discovery view
            ConfigurationPanel.Visibility = Visibility.Collapsed;
            DiscoveryPanel.Visibility = Visibility.Visible;
            StopButton.IsEnabled = true;
        }

        void StopDiscoveryButton_Click(object sender, RoutedEventArgs e)
    {
        // Stop the watcher but stay in Discovery view so the user can still interact
        // with the discovered targets.
        _watcher!.Stop();
        _watcherRunning = false;

        StopButton.IsEnabled = false;
    }

    void ReconfigureButton_Click(object sender, RoutedEventArgs e)
        {
            // Stop the watcher if we haven't already
            if (_watcherRunning)
            {
                _watcher!.Stop();
                _watcherRunning = false;
            }
            _watcher = null;

            // Clear targets from UI
            DiscoveredTargets.Clear();

            // Go back to Watcher Configuration view.
            ConfigurationPanel.Visibility = Visibility.Visible;
            DiscoveryPanel.Visibility = Visibility.Collapsed;
        }
    }
}

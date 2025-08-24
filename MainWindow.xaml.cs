using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using WinForms = System.Windows.Forms; // Alias for NotifyIcon
using NAudio.CoreAudioApi;
using NAudio.CoreAudioApi.Interfaces; // Needed for IMMNotificationClient
using System.Globalization;
using Microsoft.Win32; // Added for registry access
using NAudio.Wave;

namespace WpfMicVolumeLock
{
    public partial class MainWindow : Window, IMMNotificationClient
    {
        private readonly MMDeviceEnumerator enumerator = new();
        private MMDevice? selectedDevice;
        private float targetVolume = 1.0f;
        private WinForms.NotifyIcon? trayIcon;
        private WinForms.ContextMenuStrip? trayMenu;
        private bool isLocking = false;

        // Fields for audio testing
        private WaveInEvent? micCapture;
        private WaveOutEvent? waveOut;
        private BufferedWaveProvider? bufferedProvider;
        private bool isTesting = false;

        public MainWindow()
        {
            InitializeComponent();

            // Set the window background after resources are loaded
            this.Background = (SolidColorBrush)Resources["BackgroundBrush"];

            ApplySystemTheme(); // Apply theme before other initialization
            InitializeTrayIcon();

            // Register for device notifications
            enumerator.RegisterEndpointNotificationCallback(this);

            LoadDevices();
            LoadOutputDevices(); // Add this method to load output devices

            // Handle window state changes for tray functionality
            this.StateChanged += MainWindow_StateChanged;
            this.Closing += MainWindow_Closing;
        }

        private void ApplySystemTheme()
        {
            try
            {
                // Check Windows 10/11 theme setting
                using var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize");
                if (key?.GetValue("AppsUseLightTheme") is int lightThemeValue)
                {
                    if (lightThemeValue == 0) // Dark theme
                    {
                        SetDarkTheme();
                    }
                    else // Light theme
                    {
                        SetLightTheme();
                    }
                }
                else
                {
                    // Fallback to light theme if registry key doesn't exist
                    SetLightTheme();
                }
            }
            catch (Exception)
            {
                // If there's any error reading the registry, default to light theme
                SetLightTheme();
            }
        }

        private void SetLightTheme()
        {
            // Light theme colors - replace the existing resource brushes
            Resources["BackgroundBrush"] = new SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#F3F3F3"));
            Resources["CardBackgroundBrush"] = new SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#FFFFFF"));
            Resources["TextBrush"] = new SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#323130"));
            Resources["SubtextBrush"] = new SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#605E5C"));
            Resources["BorderBrush"] = new SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#D1D1D1"));
            Resources["StatusBackgroundBrush"] = new SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#F8F9FA"));
            Resources["StatusBorderBrush"] = new SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#E1E5E9"));
            Resources["FooterTextBrush"] = new SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#8A8886"));

            // Update window background if it was set in code
            if (this.Background is SolidColorBrush)
                this.Background = (SolidColorBrush)Resources["BackgroundBrush"];
        }

        private void SetDarkTheme()
        {
            // Dark theme colors - update the existing resource brushes
            if (Resources["BackgroundBrush"] is SolidColorBrush backgroundBrush)
                backgroundBrush.Color = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#1F1F1F");
            if (Resources["CardBackgroundBrush"] is SolidColorBrush cardBrush)
                cardBrush.Color = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#2D2D2D");
            if (Resources["TextBrush"] is SolidColorBrush textBrush)
                textBrush.Color = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#FFFFFF");
            if (Resources["SubtextBrush"] is SolidColorBrush subtextBrush)
                subtextBrush.Color = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#C8C6C4");
            if (Resources["BorderBrush"] is SolidColorBrush borderBrush)
                borderBrush.Color = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#484644");
            if (Resources["StatusBackgroundBrush"] is SolidColorBrush statusBgBrush)
                statusBgBrush.Color = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#252525");
            if (Resources["StatusBorderBrush"] is SolidColorBrush statusBorderBrush)
                statusBorderBrush.Color = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#3F3F3F");
            if (Resources["FooterTextBrush"] is SolidColorBrush footerBrush)
                footerBrush.Color = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#8A8886");
        }

        private void LoadDevices()
        {
            // Store currently selected device
            string? selectedName = comboDevices.SelectedItem as string;

            comboDevices.Items.Clear();
            var devices = enumerator.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active);

            foreach (var device in devices)
            {
                comboDevices.Items.Add(device.FriendlyName);
            }

            // Restore previous selection if still available
            if (!string.IsNullOrEmpty(selectedName) && comboDevices.Items.Contains(selectedName))
            {
                comboDevices.SelectedItem = selectedName;
            }
            else if (comboDevices.Items.Count > 0)
            {
                comboDevices.SelectedIndex = 0;
            }
        }

        private void LoadOutputDevices() // Add this method to load output devices
        {
            string? selectedName = comboOutputDevices.SelectedItem as string;

            comboOutputDevices.Items.Clear();
            var devices = enumerator.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active);

            foreach (var device in devices)
            {
                comboOutputDevices.Items.Add(device.FriendlyName);
            }

            if (!string.IsNullOrEmpty(selectedName) && comboOutputDevices.Items.Contains(selectedName))
            {
                comboOutputDevices.SelectedItem = selectedName;
            }
            else if (comboOutputDevices.Items.Count > 0)
            {
                comboOutputDevices.SelectedIndex = 0;
            }
        }

        // IMMNotificationClient implementation
        public void OnDeviceAdded(string deviceId)
        {
            Dispatcher.Invoke(LoadDevices);
        }

        public void OnDeviceRemoved(string deviceId)
        {
            Dispatcher.Invoke(LoadDevices);
        }

        public void OnDeviceStateChanged(string deviceId, DeviceState newState)
        {
            Dispatcher.Invoke(LoadDevices);
        }

        public void OnDefaultDeviceChanged(DataFlow flow, Role role, string defaultDeviceId)
        {
            // Optional: refresh if the default device changes
            Dispatcher.Invoke(LoadDevices);
        }

        public void OnPropertyValueChanged(string deviceId, PropertyKey key)
        {
            // Optional: handle property changes if needed
        }

        // Only allow digits in the TextBox
        private void VolumeTextBox_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            e.Handled = !int.TryParse(e.Text, out _);
        }

        // Handle pasting to allow only valid numbers
        private void VolumeTextBox_Pasting(object sender, DataObjectPastingEventArgs e)
        {
            if (e.DataObject.GetDataPresent(typeof(string)))
            {
                string text = (string)e.DataObject.GetData(typeof(string))!;
                if (!int.TryParse(text, out _))
                {
                    e.CancelCommand();
                }
            }
            else
            {
                e.CancelCommand();
            }
        }

        // Validate and clamp value between 0 and 100
        private void VolumeTextBox_TextChanged(object? sender, TextChangedEventArgs e)
        {
            if (sender is System.Windows.Controls.TextBox textBox)
            {
                if (int.TryParse(textBox.Text, out int value))
                {
                    if (value < 0)
                        value = 0;
                    else if (value > 100)
                        value = 100;

                    if (textBox.Text != value.ToString())
                    {
                        int caret = textBox.CaretIndex;
                        textBox.Text = value.ToString();
                        textBox.CaretIndex = Math.Min(caret, textBox.Text.Length);
                    }
                    volumeSlider.Value = value;
                }
                else if (!string.IsNullOrEmpty(textBox.Text))
                {
                    textBox.Text = "0";
                    textBox.CaretIndex = textBox.Text.Length;
                }
            }
        }

        private void InitializeTrayIcon()
        {
            trayIcon = new WinForms.NotifyIcon();

            // Load icon
            var assembly = System.Reflection.Assembly.GetExecutingAssembly();
            trayIcon.Icon = new System.Drawing.Icon(System.Windows.Application.GetResourceStream(new Uri("pack://application:,,,/favicon.ico")).Stream);
            trayIcon.Text = "Microphone Volume Lock";
            trayIcon.Visible = true;

            // Context menu setup (unchanged)
            trayMenu = new WinForms.ContextMenuStrip();
            var restoreItem = new WinForms.ToolStripMenuItem("Restore");
            restoreItem.Click += (s, e) => {
                this.Show();
                this.WindowState = WindowState.Normal;
                this.Activate();
            };
            var exitItem = new WinForms.ToolStripMenuItem("Exit");
            exitItem.Click += (s, e) => {
                this.Close();
            };
            trayMenu.Items.Add(restoreItem);
            trayMenu.Items.Add("-");
            trayMenu.Items.Add(exitItem);
            trayIcon.ContextMenuStrip = trayMenu;

            trayIcon.MouseClick += (s, e) => {
                if (e.Button == WinForms.MouseButtons.Left)
                {
                    this.Show();
                    this.WindowState = WindowState.Normal;
                    this.Activate();
                }
            };
        }

        private void TrayIcon_MouseClick(object? sender, System.Windows.Forms.MouseEventArgs e)
        {
            // Only respond to left-click
            if (e.Button == System.Windows.Forms.MouseButtons.Left)
            {
                this.Show();
                this.WindowState = WindowState.Normal;
                this.Activate();
            }
        }

        private void VolumeSlider_ValueChanged(object? sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (volumeTextBox != null)
            {
                volumeTextBox.Text = ((int)volumeSlider.Value).ToString();
            }
        }

        private void BtnStart_Click(object? sender, RoutedEventArgs e)
        {
            if (comboDevices.SelectedIndex < 0)
            {
                System.Windows.MessageBox.Show("Please select a microphone device first.", "No Device Selected",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                var devices = enumerator.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active);
                selectedDevice = devices[comboDevices.SelectedIndex];
                targetVolume = (float)(volumeSlider.Value / 100.0);

                // Set immediately
                selectedDevice.AudioEndpointVolume.MasterVolumeLevelScalar = targetVolume;

                // Subscribe to volume-change notifications
                selectedDevice.AudioEndpointVolume.OnVolumeNotification += VolumeChanged;

                isLocking = true;
                btnStart.IsEnabled = false;
                btnStop.IsEnabled = true;
                comboDevices.IsEnabled = false;
                volumeSlider.IsEnabled = false;
                volumeTextBox.IsEnabled = false;

                // Grey out the controls visually
                volumeSlider.Opacity = 0.5;
                volumeTextBox.Opacity = 0.5;

                lblStatus.Text = $"🔒 Locking '{selectedDevice.FriendlyName}' at {(int)volumeSlider.Value}%";

                if (trayIcon != null)
                    trayIcon.Text = $"Mic Volume Lock - Active ({(int)volumeSlider.Value}%)";
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Error starting volume lock: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnStop_Click(object? sender, RoutedEventArgs e)
        {
            StopLocking();
        }

        private void StopLocking()
        {
            if (selectedDevice != null)
            {
                selectedDevice.AudioEndpointVolume.OnVolumeNotification -= VolumeChanged;
                lblStatus.Text = $"🔓 Stopped locking '{selectedDevice.FriendlyName}'";
            }
            else
            {
                lblStatus.Text = "Volume locking stopped";
            }

            isLocking = false;
            btnStart.IsEnabled = true;
            btnStop.IsEnabled = false;
            comboDevices.IsEnabled = true;
            volumeSlider.IsEnabled = true;
            volumeTextBox.IsEnabled = true;

            // Restore normal appearance
            volumeSlider.Opacity = 1.0;
            volumeTextBox.Opacity = 1.0;

            if (trayIcon != null)
                trayIcon.Text = "Microphone Volume Lock - Inactive";
        }

        private void VolumeChanged(AudioVolumeNotificationData data)
        {
            if (selectedDevice != null && Math.Abs(data.MasterVolume - targetVolume) > 0.001f)
            {
                // Reset it back to target
                selectedDevice.AudioEndpointVolume.MasterVolumeLevelScalar = targetVolume;
            }
        }

        private void MainWindow_StateChanged(object? sender, EventArgs e)
        {
            if (this.WindowState == WindowState.Minimized)
            {
                this.Hide();
                trayIcon?.ShowBalloonTip(2000, "Microphone Volume Lock",
                        "Application minimized to tray", WinForms.ToolTipIcon.Info);
            }
        }

        private void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
        {
            // Stop volume locking if active
            if (isLocking)
                StopLocking();

            // Unsubscribe from volume notifications
            if (selectedDevice != null)
            {
                selectedDevice.AudioEndpointVolume.OnVolumeNotification -= VolumeChanged;
            }

            enumerator?.Dispose();

            // Dispose tray icon properly
            if (trayIcon != null)
            {
                trayIcon.Visible = false;
                trayIcon.Dispose();
            }

            // Ensure app fully shuts down
            System.Windows.Application.Current.Shutdown();
        }

        private void BtnTest_Click(object? sender, RoutedEventArgs e)
        {
            if (!isTesting)
            {
                if (comboDevices.SelectedIndex < 0)
                {
                    System.Windows.MessageBox.Show("Please select a microphone device first.", "No Device Selected",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                try
                {
                    var devices = enumerator.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active);
                    var device = devices[comboDevices.SelectedIndex];

                    micCapture = new WaveInEvent
                    {
                        DeviceNumber = comboDevices.SelectedIndex,
                        WaveFormat = new WaveFormat(44100, 1)
                    };
                    bufferedProvider = new BufferedWaveProvider(micCapture.WaveFormat);

                    micCapture.DataAvailable += (s, a) =>
                    {
                        bufferedProvider.AddSamples(a.Buffer, 0, a.BytesRecorded);
                    };

                    // Replace this line:
                    // waveOut = new WaveOutEvent();
                    // With:
                    if (comboOutputDevices.SelectedIndex >= 0)
                    {
                        var outputDevices = enumerator.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active);
                        var selectedOutputDevice = outputDevices[comboOutputDevices.SelectedIndex];
                        waveOut = new WaveOutEvent
                        {
                            DeviceNumber = comboOutputDevices.SelectedIndex
                        };
                    }
                    else
                    {
                        waveOut = new WaveOutEvent();
                    }
                    waveOut.Init(bufferedProvider);

                    micCapture.StartRecording();
                    waveOut.Play();

                    isTesting = true;
                    btnTest.Content = "🛑 Stop Test";
                    btnStart.IsEnabled = false;
                    btnStop.IsEnabled = false;
                    lblStatus.Text = $"🧪 Testing '{device.FriendlyName}' (monitoring to default output)";
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show($"Error starting test: {ex.Message}", "Error",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            else
            {
                micCapture?.StopRecording();
                micCapture?.Dispose();
                micCapture = null;

                waveOut?.Stop();
                waveOut?.Dispose();
                waveOut = null;

                bufferedProvider = null;

                isTesting = false;
                btnTest.Content = "🧪 Test";
                btnStart.IsEnabled = true;
                btnStop.IsEnabled = true;
                lblStatus.Text = "Test stopped. Select a microphone and click 'Start Locking' to begin";
            }
        }
    }
}
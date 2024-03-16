using System.Configuration;
using System.Data;
using System.Windows;
using System.Windows.Forms;
using static UsbipdGui.Usbipd;

namespace UsbipdGui
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : System.Windows.Application
    {
        private System.Windows.Forms.NotifyIcon _notifyIcon = new();
        private System.Windows.Forms.ContextMenuStrip _contextMenu = new();
        private Usbipd? _usbipd = null;
        private System.Drawing.Icon _lightThemeIcon = new System.Drawing.Icon(GetResourceStream(new Uri("resource/usbip_lighttheme.ico", UriKind.Relative)).Stream);
        private System.Drawing.Icon _darkThemeIcon = new System.Drawing.Icon(GetResourceStream(new Uri("resource/usbip_darktheme.ico", UriKind.Relative)).Stream);
        private System.Drawing.Image _bindIconImage = System.Drawing.Image.FromStream(GetResourceStream(new Uri("resource/state_bind.ico", UriKind.Relative)).Stream);
        private System.Drawing.Image _attachIconImage = System.Drawing.Image.FromStream(GetResourceStream(new Uri("resource/state_attach.ico", UriKind.Relative)).Stream);

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            _usbipd = Usbipd.BuildUsbIpdCommnad();
            if (_usbipd is null)
            {
                if (System.Windows.Forms.MessageBox.Show(
                    "'usbipd-win' is not installed.\nWould you like to visit the installation page?",
                    "usbipd-gui",
                    System.Windows.Forms.MessageBoxButtons.YesNo,
                    System.Windows.Forms.MessageBoxIcon.Asterisk
                    ) == System.Windows.Forms.DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(
                        new System.Diagnostics.ProcessStartInfo()
                        {
                            FileName = "https://github.com/dorssel/usbipd-win/releases",
                            UseShellExecute = true,
                        }
                    );
                }
                Shutdown();
                return;
            }

            _notifyIcon.Visible = true;
            _notifyIcon.Text = "usbipd";
            _notifyIcon.Icon = GetNotifyIcon();
            _notifyIcon.ContextMenuStrip = _contextMenu;
            _notifyIcon.MouseClick += new System.Windows.Forms.MouseEventHandler(ClickNotifyIcon);

            var usbDevices = _usbipd.GetUsbDevices();
            UpdateContextMenu(ref _contextMenu, ref usbDevices);

            // Add system eventt handler to switch the theme of notify icon
            Microsoft.Win32.SystemEvents.UserPreferenceChanged += SystemEvents_UserPreferenceChanged;
        }

        protected override void OnExit(ExitEventArgs e)
        {
            _contextMenu.Dispose();
            _notifyIcon.Dispose();
            base.OnExit(e);
        }

        private static bool IsLightTheme()
        {
            using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize");
            var value = key?.GetValue("AppsUseLightTheme");
            return value is int i && i > 0;
        }

        private System.Drawing.Icon GetNotifyIcon()
        {
            if (IsLightTheme()) {
                return _lightThemeIcon;
            } else {
                return _darkThemeIcon;
        }
        }

        private void SystemEvents_UserPreferenceChanged(object sender, Microsoft.Win32.UserPreferenceChangedEventArgs e)
        {
            // This event handler will be called when the system theme is changed.
            if (e.Category == Microsoft.Win32.UserPreferenceCategory.General)
            {
                // update notify icon for dark or light theme
                _notifyIcon.Icon = GetNotifyIcon();
            }
        }

        private void UpdateContextMenu(ref System.Windows.Forms.ContextMenuStrip contextMenu, ref List<UsbDevice> usbDevices)
        {
            contextMenu.Items.Clear();
            List<ToolStripMenuItem> connectedDeviceItems = [];
            List<ToolStripMenuItem> persistedDeviceItems = [];

            foreach (UsbDevice dev in usbDevices)
            {
                System.Diagnostics.Debug.WriteLine(dev);
                switch (dev.State)
                {
                    case UsbDevice.ConnectionStates.DisconnectedPersisted:
                        {
                            ToolStripMenuItem item = new ToolStripMenuItem($"none | {dev.Vid}:{dev.Pid} | {dev.Description}");
                            item.Enabled = false;
                            persistedDeviceItems.Add(item);
                        }
                        break;
                    case UsbDevice.ConnectionStates.ConnectedNotShared:
                        {
                            ToolStripMenuItem item = new ToolStripMenuItem($"{dev.BusId} | {dev.Vid}:{dev.Pid} | {dev.Description}", null, ClickUnbindedDevice);
                            item.Tag = dev;
                            connectedDeviceItems.Add(item);
                        }
                        break;
                    case UsbDevice.ConnectionStates.ConnectedShared:
                        {
                            ToolStripMenuItem item = new ToolStripMenuItem($"{dev.BusId} | {dev.Vid}:{dev.Pid} | {dev.Description} (Shared)", null, ClickBindedDevice);
                            item.Tag = dev;
                            item.Image = _bindIconImage;
                            item.Font = new System.Drawing.Font(item.Font, System.Drawing.FontStyle.Bold);
                            connectedDeviceItems.Add(item);
                        }
                        break;
                    case UsbDevice.ConnectionStates.ConnectedAttached:
                        {
                            ToolStripMenuItem item = new ToolStripMenuItem($"{dev.BusId} | {dev.Vid}:{dev.Pid} | {dev.Description} (Attached)", null, ClickAttachedDevice);
                            item.Tag = dev;
                            item.Image = _attachIconImage;
                            item.Font = new System.Drawing.Font(item.Font, System.Drawing.FontStyle.Bold);
                            connectedDeviceItems.Add(item);
                        }
                        break;
                    default:
                        break;
                }
            }

            ToolStripLabel connectedDeviceLabel = new ToolStripLabel("Connected Devices");
            connectedDeviceLabel.ForeColor = Color.Black;
            connectedDeviceLabel.Font = new System.Drawing.Font(connectedDeviceLabel.Font, System.Drawing.FontStyle.Underline);
            contextMenu.Items.Add(connectedDeviceLabel);
            contextMenu.Items.AddRange(connectedDeviceItems.ToArray());

            if (persistedDeviceItems.Count > 0)
            {
                contextMenu.Items.Add(new ToolStripSeparator());
                ToolStripLabel persistedDeviceLabel = new ToolStripLabel("Persisted Devices");
                persistedDeviceLabel.Enabled = false;
                persistedDeviceLabel.Font = new System.Drawing.Font(persistedDeviceLabel.Font, System.Drawing.FontStyle.Underline);
                contextMenu.Items.Add(persistedDeviceLabel);
                contextMenu.Items.AddRange(persistedDeviceItems.ToArray());
            }

            contextMenu.Items.Add(new ToolStripSeparator());
            contextMenu.Items.Add(new ToolStripMenuItem($"Quit", null, ClickQuit));
        }

        private void ClickNotifyIcon(object? sender, System.Windows.Forms.MouseEventArgs e)
        {
            switch (e.Button)
            {
                case MouseButtons.Right:
                    var usbDevices = _usbipd.GetUsbDevices();
                    UpdateContextMenu(ref _contextMenu, ref usbDevices);
                    break;
                default:
                    break;
            }
        }
        private void ClickUnbindedDevice(object? sender, EventArgs e)
        {
            UsbDevice device = (UsbDevice)((ToolStripMenuItem)sender).Tag;
            System.Diagnostics.Debug.WriteLine($"usbipd bind {device.BusId}");
            if (!_usbipd.Bind(ref device))
            {
                System.Diagnostics.Debug.WriteLine($"Failed to bind {device.BusId}");
            }
        }

        private void ClickBindedDevice(object? sender, EventArgs e)
        {
            UsbDevice device = (UsbDevice)((ToolStripMenuItem)sender).Tag;
            System.Diagnostics.Debug.WriteLine($"usbipd unbind {device.BusId}");
            if (!_usbipd.Unbind(ref device))
            {
                System.Diagnostics.Debug.WriteLine($"Failed to unbind {device.BusId}");
            }
        }

        private void ClickAttachedDevice(object? sender, EventArgs e)
        {
            UsbDevice device = (UsbDevice)((ToolStripMenuItem)sender).Tag;

            if (System.Windows.Forms.MessageBox.Show(
                $"\"{device.BusId} {device.Description}\" is currently attached from other machine.\nDo you really want to unbind it?",
                "usbipd-gui",
                System.Windows.Forms.MessageBoxButtons.YesNo,
                System.Windows.Forms.MessageBoxIcon.Asterisk
                ) == System.Windows.Forms.DialogResult.Yes)
            {
            System.Diagnostics.Debug.WriteLine($"usbipd unbind {device.BusId}");
            if (!_usbipd.Unbind(ref device))
            {
                System.Diagnostics.Debug.WriteLine($"Failed to unbind {device.BusId}");
            }
        }
        }

        private void ClickQuit(object? sender, EventArgs e)
        {
            Shutdown();
        }
    }

}

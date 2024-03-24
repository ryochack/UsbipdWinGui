﻿using System.Collections.Specialized;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using UsbipdGui.Properties;
using static UsbipdGui.Usbipd;

namespace UsbipdGui
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : System.Windows.Application
    {
        // Resources
        private readonly System.Drawing.Icon _lightThemeIcon = new(
            GetResourceStream(new Uri("resource/usbip_lighttheme.ico", UriKind.Relative)).Stream);
        private readonly System.Drawing.Icon _darkThemeIcon = new(
            GetResourceStream(new Uri("resource/usbip_darktheme.ico", UriKind.Relative)).Stream);
        private readonly System.Drawing.Image _bindIconImage = System.Drawing.Image.FromStream(
            GetResourceStream(new Uri("resource/state_bind.ico", UriKind.Relative)).Stream);
        private readonly System.Drawing.Image _attachIconImage = System.Drawing.Image.FromStream(
            GetResourceStream(new Uri("resource/state_attach.ico", UriKind.Relative)).Stream);

        private Usbipd? _usbipd = null;
        Settings _settings = new();
        private List<UsbDevice> _ignoredDeviceList = [];
        private System.Windows.Forms.NotifyIcon _notifyIcon = new();
        private System.Windows.Forms.ContextMenuStrip _contextMenu = new();

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
            _notifyIcon.MouseClick += new System.Windows.Forms.MouseEventHandler(OnClickToShowAppMenu);

            _ignoredDeviceList = LoadIgnoredUsbIdList();

            // Add system eventt handler to switch the theme of notify icon
            Microsoft.Win32.SystemEvents.UserPreferenceChanged += SystemEvents_UserPreferenceChanged;
        }

        private List<UsbDevice> LoadIgnoredUsbIdList()
        {
            Regex regex = UsbIdRegex();
            List<UsbDevice> usbDevices = [];

            List<string> listAsString = _settings.IgnoredUsbIds?.Cast<string>().ToList() ?? [];
            Debug.WriteLine("===== Load Ignored Usb ID List =====");
            foreach (string s in listAsString)
            {
                Debug.WriteLine(s);
                Match match = regex.Match(s);
                if (match.Success)
                {
                    usbDevices.Add(new UsbDevice(match.Groups[1].Value, match.Groups[2].Value));
                }
            }
            Debug.WriteLine("==========");
            return usbDevices;
        }

        // Regex target examples: "8087:0025"
        [GeneratedRegex(@"(\w{4}):(\w{4})")]
        private static partial Regex UsbIdRegex();

        private void SaveIgnoredUsbIdList(in List<UsbDevice> usbDevices)
        {
            _settings.IgnoredUsbIds.Clear();
            _settings.IgnoredUsbIds.AddRange(usbDevices.Select(dev => $"{dev.Vid}:{dev.Pid}").ToArray());
            _settings.Save();
        }

        protected override void OnExit(ExitEventArgs e)
        {
            _contextMenu.Dispose();
            _notifyIcon.Dispose();
            base.OnExit(e);
        }

        private static bool IsLightTheme()
        {
            using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize");
            var value = key?.GetValue("AppsUseLightTheme");
            return value is int i && i > 0;
        }

        private System.Drawing.Icon GetNotifyIcon()
        {
            if (IsLightTheme())
            {
                return _lightThemeIcon;
            }
            else
            {
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

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private ToolStripMenuItem BuildDisconnectedPersistedDeviceItem(in UsbDevice dev)
        {
            ToolStripMenuItem item = new($"none | {dev.Vid}:{dev.Pid} | {dev.Description}")
            {
                Enabled = false
            };
            return item;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private ToolStripMenuItem BuildConnectedNotSharedDeviceItem(in UsbDevice dev)
        {
            string desc = $"{dev.BusId} | {dev.Vid}:{dev.Pid} | {dev.Description}";
            ToolStripMenuItem item = new(desc)
            {
                Tag = dev,
                ToolTipText = "Bind this device (or Right click to ignore)"
            };
            item.MouseUp += (sender, e) =>
            {
                if (e.Button == MouseButtons.Left)
                {
                    OnLeftClickToBindDevice(sender, e);
                }
                else if (e.Button == MouseButtons.Right)
                {
                    if ((sender as ToolStripMenuItem)?.Tag is not UsbDevice device) { return; }
                    ContextMenuStrip subMenu = new();
                    ToolStripMenuItem item = new($"Ignore {desc}", null, OnLeftClickToAddIgnoreList)
                    {
                        Tag = device
                    };
                    subMenu.Items.Add(item);
                    subMenu.Show(Cursor.Position);
                }
            };
            return item;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private ToolStripMenuItem BuildConnectedSharedDeviceItem(in UsbDevice dev)
        {
            ToolStripMenuItem item = new($"{dev.BusId} | {dev.Vid}:{dev.Pid} | {dev.Description} (Shared)")
            {
                Tag = dev,
                Image = _bindIconImage,
                ToolTipText = "Unbind this device",
            };
            item.Font = new System.Drawing.Font(item.Font, System.Drawing.FontStyle.Bold);
            item.MouseUp += (sender, e) =>
            {
                if (e.Button == MouseButtons.Left)
                {
                    OnLeftClickToUnbindDevice(sender, e);
                }
            };
            return item;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private ToolStripMenuItem BuildConnectedAttachedDeviceItem(in UsbDevice dev)
        {
            ToolStripMenuItem item = new(
                $"{dev.BusId} | {dev.Vid}:{dev.Pid} | {dev.Description} (Attached from {dev.ClientIpAddr})")
            {
                Tag = dev,
                Image = _attachIconImage,
                ToolTipText = "Unbind this device"
            };
            item.Font = new System.Drawing.Font(item.Font, System.Drawing.FontStyle.Bold);
            item.MouseUp += (sender, e) =>
            {
                if (e.Button == MouseButtons.Left)
                {
                    OnLeftClickToUnbindDeviceWithCaution(sender, e);
                }
            };
            return item;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private ToolStripMenuItem BuildIgnoredDeviceItem(in UsbDevice dev, in string toolTips)
        {
            ToolStripMenuItem item = new($"{dev.BusId ?? "none"} | {dev.Vid}:{dev.Pid} | {dev.Description}")
            {
                Tag = dev,
                ToolTipText = toolTips
            };
            item.MouseUp += (sender, e) =>
            {
                if (e.Button == MouseButtons.Left)
                {
                    OnLeftClickToRemoveFromIgnoreList(sender, e);
                }
            };
            return item;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private ToolStripMenuItem BuildQuitItem()
        {
            ToolStripMenuItem item = new($"Quit");
            item.MouseUp += (sender, e) =>
            {
                if (e.Button == MouseButtons.Left)
                {
                    OnLeftClickToQuit(sender, e);
                }
            };
            return item;
        }

        private void UpdateContextMenu(ref System.Windows.Forms.ContextMenuStrip contextMenu, in List<UsbDevice> usbDevices)
        {
            contextMenu.Items.Clear();
            List<ToolStripMenuItem> connectedDeviceItems = [];
            List<ToolStripMenuItem> persistedDeviceItems = [];

            // Update ignored list
            for (int i = 0; i < _ignoredDeviceList.Count; i++)
            {
                UsbDevice? matchedDevice = usbDevices.FirstOrDefault(dev => ((dev.Vid == _ignoredDeviceList[i].Vid) && (dev.Pid == _ignoredDeviceList[i].Pid)));
                if (matchedDevice is null) {
                    continue;
                }
                if ((matchedDevice.Vid is not null) && (matchedDevice.Pid is not null))
                {
                    _ignoredDeviceList[i] = matchedDevice;
                }
                else
                {
                    _ignoredDeviceList[i] = new UsbDevice(_ignoredDeviceList[i].Description, _ignoredDeviceList[i].Vid, _ignoredDeviceList[i].Pid);
                }
            }

            foreach (UsbDevice dev in usbDevices.Except(_ignoredDeviceList, new UsbIdEqualityComparer()))
            {
                System.Diagnostics.Debug.WriteLine(dev);
                switch (dev.State)
                {
                    case UsbDevice.ConnectionStates.DisconnectedPersisted:
                        persistedDeviceItems.Add(BuildDisconnectedPersistedDeviceItem(dev));
                        break;
                    case UsbDevice.ConnectionStates.ConnectedNotShared:
                        connectedDeviceItems.Add(BuildConnectedNotSharedDeviceItem(dev));
                        break;
                    case UsbDevice.ConnectionStates.ConnectedShared:
                        connectedDeviceItems.Add(BuildConnectedSharedDeviceItem(dev));
                        break;
                    case UsbDevice.ConnectionStates.ConnectedAttached:
                        connectedDeviceItems.Add(BuildConnectedAttachedDeviceItem(dev));
                        break;
                    default:
                        break;
                }
            }

            // List connected devices
            {
                ToolStripLabel connectedDeviceLabel = new("Connected USB Devices");
                connectedDeviceLabel.Font = new System.Drawing.Font(connectedDeviceLabel.Font, System.Drawing.FontStyle.Underline);
                contextMenu.Items.Add(connectedDeviceLabel);
                contextMenu.Items.AddRange(connectedDeviceItems.ToArray());
            }

            // List persisted devices
            if (persistedDeviceItems.Count > 0)
            {
                contextMenu.Items.Add(new ToolStripSeparator());
                ToolStripLabel persistedDeviceLabel = new("Persisted USB Devices")
                {
                    Enabled = false
                };
                persistedDeviceLabel.Font = new System.Drawing.Font(persistedDeviceLabel.Font, System.Drawing.FontStyle.Underline);
                contextMenu.Items.Add(persistedDeviceLabel);
                contextMenu.Items.AddRange(persistedDeviceItems.ToArray());
            }

            // Drop Down item of ignored devices
            {
                contextMenu.Items.Add(new ToolStripSeparator());
                ToolStripMenuItem dropDownMenu = new("Ignored USB Device List...");
                if (_ignoredDeviceList.Count == 0)
                {
                    dropDownMenu.Enabled = false;
                }
                else
                {
                    List<ToolStripMenuItem> connectedIgnoreList = [];
                    List<ToolStripMenuItem> disconnectedIgnoreList = [];
                    foreach (UsbDevice dev in _ignoredDeviceList)
                    {
                        if ((dev.State & UsbDevice.ConnectionStates.Connected) != 0)
                        {
                            connectedIgnoreList.Add(BuildIgnoredDeviceItem(dev, "Restore from ignored list"));
                        }
                        else
                        {
                            disconnectedIgnoreList.Add(BuildIgnoredDeviceItem(dev, "Remove from ignored list"));
                        }
                    }
                    if (connectedIgnoreList.Count > 0)
                    {
                        ToolStripLabel connectedDeviceLabel = new("Connected USB Devices");
                        connectedDeviceLabel.Font = new System.Drawing.Font(connectedDeviceLabel.Font, System.Drawing.FontStyle.Underline);
                        dropDownMenu.DropDownItems.Add(connectedDeviceLabel);
                        dropDownMenu.DropDownItems.AddRange(connectedIgnoreList.ToArray());
                    }
                    if (disconnectedIgnoreList.Count > 0)
                    {
                        ToolStripLabel disconnectedDeviceLabel = new("Disconnected USB Devices");
                        disconnectedDeviceLabel.Font = new System.Drawing.Font(disconnectedDeviceLabel.Font, System.Drawing.FontStyle.Underline);
                        dropDownMenu.DropDownItems.Add(disconnectedDeviceLabel);
                        dropDownMenu.DropDownItems.AddRange(disconnectedIgnoreList.ToArray());
                    }
                }
                contextMenu.Items.Add(dropDownMenu);
            }

            // Quit item
            contextMenu.Items.Add(new ToolStripSeparator());
            contextMenu.Items.Add(BuildQuitItem());
        }

        private void OnClickToShowAppMenu(object? sender, System.Windows.Forms.MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right && _usbipd is not null)
            {
                UpdateContextMenu(ref _contextMenu, _usbipd.GetUsbDevices());
                _contextMenu.Show();
            }
        }

        private void OnLeftClickToAddIgnoreList(object? sender, EventArgs e)
        {
            if ((sender as ToolStripMenuItem)?.Tag is not UsbDevice device) {
                return;
            }
            System.Diagnostics.Debug.WriteLine($"Ignore => {device.Vid}:{device.Pid} {device.Description}");
            _ignoredDeviceList.Add(device);
            SaveIgnoredUsbIdList(_ignoredDeviceList);
        }

        private void OnLeftClickToRemoveFromIgnoreList(object? sender, EventArgs e)
        {
            if ((sender as ToolStripMenuItem)?.Tag is not UsbDevice device) {
                return;
            }
            System.Diagnostics.Debug.WriteLine($"Unignore => {device.Vid}:{device.Pid} {device.Description}");
            _ignoredDeviceList.Remove(device);
            SaveIgnoredUsbIdList(_ignoredDeviceList);
        }

        private void OnLeftClickToBindDevice(object? sender, EventArgs e)
        {
            if ((sender as ToolStripMenuItem)?.Tag is not UsbDevice device) {
                return;
            }
            System.Diagnostics.Debug.WriteLine($"usbipd bind {device.BusId}");
            if (!_usbipd?.Bind(ref device) ?? false)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to bind {device.BusId}");
            }
        }

        private void OnLeftClickToUnbindDevice(object? sender, EventArgs e)
        {
            if ((sender as ToolStripMenuItem)?.Tag is not UsbDevice device) {
                return;
            }
            System.Diagnostics.Debug.WriteLine($"usbipd unbind {device.BusId}");
            if (!_usbipd?.Unbind(ref device) ?? false)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to unbind {device.BusId}");
            }
        }

        private void OnLeftClickToUnbindDeviceWithCaution(object? sender, EventArgs e)
        {
            if ((sender as ToolStripMenuItem)?.Tag is not UsbDevice device) {
                return;
            }
            if (System.Windows.Forms.MessageBox.Show(
                $"\"{device.BusId} {device.Description}\" is currently attached from {device.ClientIpAddr}.\nDo you really want to unbind it?",
                "usbipd-gui",
                System.Windows.Forms.MessageBoxButtons.YesNo,
                System.Windows.Forms.MessageBoxIcon.Asterisk
                ) == System.Windows.Forms.DialogResult.Yes)
            {
                System.Diagnostics.Debug.WriteLine($"usbipd unbind {device.BusId}");
                if (!_usbipd?.Unbind(ref device) ?? false)
                {
                    System.Diagnostics.Debug.WriteLine($"Failed to unbind {device.BusId}");
                }
            }
        }

        private void OnLeftClickToQuit(object? sender, EventArgs e)
        {
            Shutdown();
        }
    }

}

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
            _notifyIcon.Icon = new System.Drawing.Icon(GetResourceStream(new Uri("resource/usbip_darktheme.ico", UriKind.Relative)).Stream);
            _notifyIcon.ContextMenuStrip = _contextMenu;
            _notifyIcon.MouseClick += new System.Windows.Forms.MouseEventHandler(ClickNotifyIcon);

            var usbDevices = _usbipd.GetUsbDevices();
            UpdateContextMenu(ref _contextMenu, ref usbDevices);
        }

        private void UpdateContextMenu(ref System.Windows.Forms.ContextMenuStrip contextMenu, ref List<UsbDevice> usbDevices)
        {
            contextMenu.Items.Clear();
            contextMenu.Text = "hogehoge";
            List<UsbDevice> persistedDevices = [];

            // TODO: ConnectedDevice のタイトルを入れたい
            foreach (UsbDevice dev in usbDevices)
            {
                System.Diagnostics.Debug.WriteLine(dev);
                switch (dev.State)
                {
                    case UsbDevice.ConnectionStates.DisconnectedPersisted:
                        persistedDevices.Add(dev);
                        break;
                    case UsbDevice.ConnectionStates.ConnectedNotShared:
                        contextMenu.Items.Add($"[{dev.BusId}] {dev.Vid}:{dev.Pid} {dev.Description}", null, ClickBindedDevice);
                        ((ToolStripMenuItem)contextMenu.Items[contextMenu.Items.Count - 1]).Tag = dev;
                        break;
                    case UsbDevice.ConnectionStates.ConnectedShared:
                        contextMenu.Items.Add($"[{dev.BusId}] {dev.Vid}:{dev.Pid} {dev.Description}", null, ClickUnbindedDevice);
                        ((ToolStripMenuItem)contextMenu.Items[contextMenu.Items.Count - 1]).Checked = true;
                        ((ToolStripMenuItem)contextMenu.Items[contextMenu.Items.Count - 1]).Tag = dev;
                        break;
                    case UsbDevice.ConnectionStates.ConnectedAttached:
                        contextMenu.Items.Add($"[{dev.BusId}] {dev.Vid}:{dev.Pid} {dev.Description}");
                        ((ToolStripMenuItem)contextMenu.Items[contextMenu.Items.Count - 1]).Checked = true;
                        ((ToolStripMenuItem)contextMenu.Items[contextMenu.Items.Count - 1]).Tag = dev;
                        break;
                    default:
                        break;
                }
            }
            // TODO: PersisteDevice のタイトルを入れたい
            foreach (UsbDevice dev in persistedDevices)
            {
                contextMenu.Items.Add($"[   ] {dev.Vid}:{dev.Pid} {dev.Description}");
                contextMenu.Items[contextMenu.Items.Count - 1].Enabled = false;
            }
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
        private void ClickBindedDevice(object? sender, EventArgs e)
        {
            UsbDevice device = (UsbDevice)((ToolStripMenuItem)sender).Tag;
            System.Diagnostics.Debug.WriteLine($"usbipd bind {device.BusId}");
        }

        private void ClickUnbindedDevice(object? sender, EventArgs e)
        {
            UsbDevice device = (UsbDevice)((ToolStripMenuItem)sender).Tag;
            System.Diagnostics.Debug.WriteLine($"usbipd unbind {device.BusId}");
        }

    }

}

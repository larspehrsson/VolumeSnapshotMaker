using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Principal;
using System.Windows;
using System.Windows.Forms;
using Microsoft.Win32.TaskScheduler;
using MessageBox = System.Windows.MessageBox;

namespace SnapshotMaker
{
    public class vssStorage
    {
        public string volume { get; set; }
        public string used { get; set; }
        public string allocated { get; set; }
        public string maximum { get; set; }
        public int number { get; set; }
        public DateTime oldest { get; set; }
    }

    public class shadow
    {
        public string volume { get; set; }
        public DateTime created { get; set; }
    }

    /// <summary>
    ///     Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private const string taskname = "Update VSS previous versions task";
        private static readonly Timer RunTimer = new Timer();
        private static readonly NotifyIcon Notification = new NotifyIcon();
        private List<string> SelectedDrivesList = new List<string>();
        private List<vssStorage> VSSStorageList = new List<vssStorage>();
        private List<shadow> shadowList = new List<shadow>();

        public MainWindow()
        {
            InitializeComponent();

            if (!IsElevated)
            {
                MessageBox.Show("Need to run as administrator. ", "Run as admin",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                Close();
                return;
            }

            var drivesList = DriveInfo.GetDrives().ToList();
            var minutes = 120;

            if (!string.IsNullOrEmpty(settings.saveinterval))
            {
                minutes = int.Parse(settings.saveinterval);
                SelectedDrivesList = settings.drivelist.Split('!').ToList();
            }

            HourInterval.Text = (minutes / 60).ToString();
            MinuteInterval.Text = (minutes % 60).ToString();

            DrivesListBox.ItemsSource = drivesList;
            NotificationCheckBox.IsChecked = settings.notification;

            foreach (var f in DrivesListBox.Items)
                if (SelectedDrivesList.Any(c => c == f.ToString()))
                    DrivesListBox.SelectedItems.Add(f);

            // Initialize menuExit
            var menuExit = new MenuItem
            {
                Index = 0,
                Text = "E&xit"
            };
            menuExit.Click += MenuExitClick;

            // Initialize menuRun
            var menuRun = new MenuItem
            {
                Index = 1,
                Text = "&Run"
            };
            menuRun.Click += MenuRunClick;

            // Initialize menuOpen
            var menuOpen = new MenuItem
            {
                Index = 1,
                Text = "&Open"
            };
            menuOpen.Click += MenuOpenClick;

            var contextMenu = new ContextMenu();
            contextMenu.MenuItems.Add(menuExit);
            contextMenu.MenuItems.Add(menuRun);
            contextMenu.MenuItems.Add(menuOpen);

            //System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
            //string[] names = assembly.GetManifestResourceNames();
            //foreach (string name in names)
            //    Debug.WriteLine(name);

            // The ContextMenu property sets the menu that will appear when the systray icon is right clicked.
            Notification.ContextMenu = contextMenu;
            Notification.Icon =
                new Icon(Assembly.GetExecutingAssembly().GetManifestResourceStream("SnapshotMaker.icon.ico"));

            Notification.DoubleClick +=
                delegate
                {
                    Show();
                };

            // Run every "minutes"
            RunTimer.Tick += RunVSS;
            RunTimer.Interval = 1000 * 60 * minutes;
            RunTimer.Start();

            Notification.Visible = true;

            getShadowStorage();
            GetShadows();

            if (SelectedDrivesList.Count > 0)
                Hide();

            RunVSS(null, null);
        }

        /// <summary>
        /// minimizes to systemtray, unless you have rightclicked on tray icon and selected "exit"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Window_Closing(object sender, CancelEventArgs e)
        {
            if (!forceexit)
            {
                e.Cancel = true;
                Hide();
            }
        }

        /// <summary>
        ///     Check that the program is running with elevated rights
        /// </summary>
        private bool IsElevated =>
            new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator);

        private bool forceexit = false;

        /// <summary>
        /// Closes the program
        /// </summary>
        /// <param name="Sender"></param>
        /// <param name="e"></param>
        private void MenuExitClick(object Sender, EventArgs e)
        {
            // Close the form, which closes the application.
            Notification.Visible = false;
            forceexit = true;
            Close();
        }

        /// <summary>
        /// Create a new volume shadow copy
        /// </summary>
        /// <param name="Sender"></param>
        /// <param name="e"></param>
        private async void MenuRunClick(object Sender, EventArgs e)
        {
            RunVSS(null, null);
        }

        /// <summary>
        /// Open the main window
        /// </summary>
        /// <param name="Sender"></param>
        /// <param name="e"></param>
        private async void MenuOpenClick(object Sender, EventArgs e)
        {
            getShadowStorage();
            GetShadows();
            Show();
        }

        /// <summary>
        /// Creates the VSS history
        /// </summary>
        /// <param name="myObject"></param>
        /// <param name="myEventArgs"></param>
        private void RunVSS(object myObject, EventArgs myEventArgs)
        {
            if (!IsElevated)
                return;

            // wmic shadowcopy call create Volume=C:\

            var error = "";
            var lastresult = "";

            foreach (var drive in SelectedDrivesList)
            {
                var proc = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = "wmic",
                        Arguments = $@"shadowcopy call create Volume={drive}",
                        UseShellExecute = false,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        CreateNoWindow = true,
                        Verb = "runas"
                    }
                };

                proc.Start();
                while (!proc.StandardOutput.EndOfStream)
                {
                    var line = proc.StandardOutput.ReadLine();
                    Debug.WriteLine(line);
                    lastresult += line + Environment.NewLine;
                }

                while (!proc.StandardError.EndOfStream)
                {
                    var line = proc.StandardError.ReadLine();
                    Debug.WriteLine(line);
                    if (line != "")
                        error += line + Environment.NewLine;
                    lastresult += line + Environment.NewLine;
                }
            }

            if (error != "")
            {
                Notification.BalloonTipText = $"Error updating vss {error}";
                Notification.ShowBalloonTip(10);
            }
            else
            {
                if (settings.notification)
                {
                    Notification.BalloonTipText = "VSS update successfull";
                    Notification.ShowBalloonTip(5);
                }
            }

            Notification.Text = $"Last update at {DateTime.Now} {(error != "" ? "FAILED" : "was successful")} ";
        }

        private void getShadowStorage()
        {
            if (!IsElevated)
                return;

            // vssadmin list shadowstorage
            VSSStorageList = new List<vssStorage>();

            var error = "";
            var proc = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = "vssadmin",
                    Arguments = "list shadowstorage",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                    Verb = "runas"
                }
            };

            vssStorage vss = new vssStorage();
            proc.Start();
            while (!proc.StandardOutput.EndOfStream)
            {
                var line = proc.StandardOutput.ReadLine();
                Debug.WriteLine(line);
                var kolidx = line.IndexOf(":");
                if (line.Contains("For volume"))
                {
                    vss.volume = line.Substring(kolidx + 3, 2);
                }
                if (line.Contains("  Used"))
                {
                    vss.used = line.Substring(kolidx + 2);
                }
                if (line.Contains("  Allocated"))
                {
                    vss.allocated = line.Substring(kolidx + 2);
                }
                if (line.Contains("  Maximum"))
                {
                    vss.maximum = line.Substring(kolidx + 2);
                }

                if (line.Trim() == "" && !string.IsNullOrEmpty(vss.volume))
                {
                    VSSStorageList.Add(vss);
                    vss = new vssStorage();
                }
            }

            while (!proc.StandardError.EndOfStream)
            {
                var line = proc.StandardError.ReadLine();
                Debug.WriteLine(line);
                if (line != "")
                    error += line + Environment.NewLine;
            }

            if (error != "")
            {
                Notification.BalloonTipText = $"Error getting storage usage {error}";
                Notification.ShowBalloonTip(10);
            }

            UsageListBox.ItemsSource = null;
            UsageListBox.ItemsSource = VSSStorageList.OrderBy(c => c.volume);
        }

        private void GetShadows()
        {
            if (!IsElevated)
                return;

            // vssadmin list shadows

            var error = "";

            var proc = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = "vssadmin",
                    Arguments = "list shadows",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                    Verb = "runas"
                }
            };

            shadow shadow = new shadow();
            proc.Start();
            while (!proc.StandardOutput.EndOfStream)
            {
                var line = proc.StandardOutput.ReadLine();
                Debug.WriteLine(line);
                var kolidx = line.IndexOf(":");
                if (line.Contains("  Original Volume"))
                {
                    shadow.volume = line.Substring(kolidx + 3, 2);
                }
                if (line.Contains(" creation time:"))
                {
                    var substr = line.Substring(kolidx + 1);
                    shadow.created = DateTime.Parse(substr);
                }

                if (line.Trim() == "" && !string.IsNullOrEmpty(shadow.volume))
                {
                    shadowList.Add(shadow);
                    shadow = new shadow();
                }
            }

            while (!proc.StandardError.EndOfStream)
            {
                var line = proc.StandardError.ReadLine();
                Debug.WriteLine(line);
                if (line != "")
                    error += line + Environment.NewLine;
            }

            if (error != "")
            {
                Notification.BalloonTipText = $"Error getting storage usage {error}";
                Notification.ShowBalloonTip(10);
            }

            foreach (var vss in VSSStorageList)
            {
                var shadows = shadowList.Where(c => c.volume == vss.volume).ToList();
                if (vss != null)
                {
                    vss.number = shadows.Count;
                    vss.oldest = shadows.Min(c => c.created);
                }
            }
        }

        /// <summary>
        /// Saves the configuration to the registry
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            var hours = int.Parse(HourInterval.Text);
            var minutes = int.Parse(MinuteInterval.Text);
            SelectedDrivesList = DrivesListBox.SelectedItems.Cast<DriveInfo>().ToList().Select(c => c.ToString())
                .ToList();

            settings.drivelist = string.Join("!", SelectedDrivesList.ToArray());
            settings.saveinterval = (hours * 60 + minutes).ToString();
            settings.notification = NotificationCheckBox.IsChecked.Value;
            settings.Save();

            if (!isTaskInstalled())
                CreateTask();

            Hide();

            RunVSS(null, null);
        }

        /// <summary>
        /// Forces a run of VSS
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RunButton_Click(object sender, RoutedEventArgs e)
        {
            SelectedDrivesList = DrivesListBox.SelectedItems.Cast<DriveInfo>().ToList().Select(c => c.ToString())
                .ToList();

            RunVSS(null, null);
        }

        /// <summary>
        /// Checks to see if the task is already installed in Task Scheduler
        /// </summary>
        /// <returns></returns>
        private bool isTaskInstalled()
        {
            return TaskService.Instance.GetTask($@"\{taskname}") != null;
        }

        /// <summary>
        /// Creates a task in task-scheduler. Runs the program in elevated mode
        /// </summary>
        private void CreateTask()
        {
            if (!IsElevated)
                return;

            // Create a new task definition for the local machine and assign properties
            var td = TaskService.Instance.NewTask();
            td.RegistrationInfo.Description = "Update VSS previous versions";

            var lt = new LogonTrigger
            {
                UserId = WindowsIdentity.GetCurrent().Name
            };
            td.Triggers.Add(lt);

            // Create an action that will launch Notepad whenever the trigger fires
            td.Actions.Add(
                Process.GetCurrentProcess().MainModule.FileName,
                "",
                Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName));

            td.Principal.RunLevel = TaskRunLevel.Highest;
            td.RegistrationInfo.Author = WindowsIdentity.GetCurrent().Name;

            // Register the task in the root folder of the local machine
            TaskService.Instance.RootFolder.RegisterTaskDefinition(taskname, td);
        }

        /// <summary>
        /// Removes the task from the Task Scheduler
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RemoveTask_Click(object sender, RoutedEventArgs e)
        {
            if (!isTaskInstalled())
                return;

            TaskService.Instance.RootFolder.DeleteTask($@"\{taskname}");

            MessageBox.Show("Task has been removed ", "Remove task",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }

        private void ButtonBase_OnClick(object sender, RoutedEventArgs e)
        {
            var proc = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = "C:\\Windows\\System32\\SystemPropertiesProtection.exe",
                    Arguments = "",
                    //UseShellExecute = false,
                    //RedirectStandardOutput = true,
                    //RedirectStandardError = true,
                    //CreateNoWindow = true,
                    Verb = "runas"
                }
            };

            proc.Start();
        }
    }
}
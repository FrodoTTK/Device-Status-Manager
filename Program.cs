using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using System.Diagnostics;
using System.Drawing;
using DeviceStatusManager.Common;
using DeviceStatusManager.Entitys;
using System.Linq;
using NLog;
using NLog.Config;
using NLog.Targets;
using IWshRuntimeLibrary;
using System.Collections;
using Newtonsoft.Json;
using System.Runtime.Remoting.Lifetime;
using static DeviceStatusManager.Common.CommonFunctions;

namespace DeviceManager
{
    static class Program
    {
        static Dictionary<string, List<string>> DeviceStates { get; set; }
        static NotifyIcon notifyIcon;
        static Dictionary<string, Dictionary<string, List<string>>> deviceStatesColl = new Dictionary<string, Dictionary<string, List<string>>>();
        static List<string> currentStates = new List<string>();
        static Device[] devices;
        static string cDir = Directory.GetCurrentDirectory();
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        static string startupFolder = Environment.GetFolderPath(Environment.SpecialFolder.Startup);

        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            #region Initialization LogManager
            // create logs floder
            var config = new LoggingConfiguration();
            Directory.CreateDirectory("logs");

            var fileTarget = new FileTarget("logfile")
            {
                FileName = "logs/${shortdate}.txt", // log file name with date
                Layout = "${longdate} ${uppercase:[${level}]}: ${message} ${exception}", // log template
            };

            // configure LogManager
            config.AddTarget(fileTarget);
            config.AddRule(LogLevel.Info, LogLevel.Fatal, fileTarget);
            LogManager.Configuration = config;
            #endregion

            ClearLogFiles(logger);

            // load states from 'config.json'
            ConfigOperator.LoadConfig(logger, ref deviceStatesColl);

            notifyIcon = new NotifyIcon
            {
                Text= "Device Status Manager",
                Icon = SystemIcons.Application,
                //Icon = new Icon(Directory.GetCurrentDirectory() + "\\resources\\icons\\appicon.ico"),
                Visible = true,
                ContextMenuStrip = new ContextMenuStrip()
            };
            notifyIcon.ContextMenuStrip.Renderer = new CustomToolStripRenderer();

            int collIndex = 0;
            foreach (var deviceStates in deviceStatesColl)
            {
                var allDevices = deviceStates.Value.SelectMany(d => d.Value).ToList();
                var allDevicesDistinct = allDevices.Distinct();
                deviceStates.Value.Add("EnableAll", allDevicesDistinct.ToList());

                // load all PnP devices information
                ConfigOperator.ReloadDevicesInfo(logger, cDir, ref devices);

                currentStates.Add("EnableAll");
                UpdateContextMenu(deviceStates.Key, deviceStates.Value, collIndex);
                collIndex++;
            }

            // apply last exit state
            if (System.IO.File.Exists("lastState"))
            {
                List<string> lastState = JsonConvert.DeserializeObject<List<string>>(System.IO.File.ReadAllText("lastState"));
                for (int i = 0; i < lastState.Count(); i++)
                {
                    ApplyState(lastState[i], i);
                }
            }
            else
            {
                FileStream fs = System.IO.File.Create("lastState");
                fs.Close();
                for (int i = 0; i < deviceStatesColl.Count(); i++)
                {
                    ApplyState("EnableAll", i);
                }
            }

            AddOperationButton();

            logger.Info("Starting application...");
            Application.Run();
        }
        // update menu
        static void UpdateContextMenu(string collName, Dictionary<string, List<string>> deviceStates, int collIndex)
        {
            ToolStripMenuItem menuItem = new ToolStripMenuItem(collName) { Enabled = false, Checked = false };
            notifyIcon.ContextMenuStrip.Items.Add(menuItem);

            foreach (var state in deviceStates.Keys)
            {
                menuItem = new ToolStripMenuItem(state, null, (sender, e) => OnMenuItemClick(sender, e, collIndex))
                {
                    Checked = state == currentStates[collIndex]
                };
                notifyIcon.ContextMenuStrip.Items.Add(menuItem);
            }
            ToolStripSeparator separator = new ToolStripSeparator();
            notifyIcon.ContextMenuStrip.Items.Add(separator);
        }

        static void OnMenuItemClick(object sender, EventArgs e, int collIndex)
        {
            var menuItem = sender as ToolStripMenuItem;
            ConfigOperator.ReloadDevicesInfo(logger, cDir, ref devices);
            ApplyState(menuItem.Text, collIndex);
            ReloadConfig();
        }
        static void OnMenuItemClick(object sender, EventArgs e)
        {
            var menuItem = sender as ToolStripMenuItem;
            switch (menuItem.Text)
            {
                case "Start at login":
                    if (menuItem.Checked)
                    {
                        System.IO.File.Delete(startupFolder + @"\DeviceStatusManager.lnk");
                        menuItem.Checked = false;
                    }
                    else
                    {
                        // create a shortcut on startup
                        IWshShortcut shortcut = StartAtLogin(startupFolder);
                        shortcut.WorkingDirectory = Application.StartupPath; // working directory
                        shortcut.TargetPath = Application.ExecutablePath; // path of the executable
                        shortcut.Save();
                        menuItem.Checked = true;
                    }
                    break;
                case "Open directory":
                    Process.Start("explorer.exe", Directory.GetCurrentDirectory());
                    break;
                case "Reload config":
                    ConfigOperator.LoadConfig(logger, ref deviceStatesColl);
                    ReloadConfig();
                    break;
                case "Exit":
                    Application.Exit();
                    break;
                default:
                    break;
            }
        }
        static void ReloadConfig()
        {
            ConfigOperator.LoadConfig(logger, ref deviceStatesColl);
            notifyIcon.ContextMenuStrip.Items.Clear();
            int collIndex = 0;
            foreach (var deviceStates in deviceStatesColl)
            {
                var allDevices = deviceStates.Value.SelectMany(d => d.Value).ToList();
                var allDevicesDistinct = allDevices.Distinct();
                deviceStates.Value.Add("EnableAll", allDevicesDistinct.ToList());
                UpdateContextMenu(deviceStates.Key, deviceStates.Value, collIndex);
                collIndex++;
            }
            AddOperationButton();
        }
        static void AddOperationButton()
        {
            ToolStripMenuItem menuItem = new ToolStripMenuItem("Start at login", null, (sender, e) => OnMenuItemClick(sender, e)) { Checked = System.IO.File.Exists(startupFolder + @"\DeviceStatusManager.lnk") };
            notifyIcon.ContextMenuStrip.Items.Add(menuItem);
            menuItem = new ToolStripMenuItem("Open directory", null, (sender, e) => OnMenuItemClick(sender, e)) { Checked = false };
            notifyIcon.ContextMenuStrip.Items.Add(menuItem);
            menuItem = new ToolStripMenuItem("Reload config", null, (sender, e) => OnMenuItemClick(sender, e)) { Checked = false };
            notifyIcon.ContextMenuStrip.Items.Add(menuItem);
            menuItem = new ToolStripMenuItem("Exit", null, (sender, e) => OnMenuItemClick(sender, e)) { Checked = false };
            notifyIcon.ContextMenuStrip.Items.Add(menuItem);
        }
        static void ApplyState(string state, int collIndex)
        {
            logger.Info($"Staring handle devices of state: '{state}'.");
            List<string> collKeys = new List<string>(deviceStatesColl.Keys);
            var coll = deviceStatesColl[collKeys[collIndex]];
            if (state == "EnableAll")
            {
                TraversalDevices(state, coll.First(s => s.Key == state));
            }
            else
            {
                foreach (var stateItem in coll)
                {
                    if (!stateItem.Value.Any())
                    {
                        logger.Info($"state: '{state}' not configure any devices.");
                        continue;
                    }
                    if (stateItem.Key == "EnableAll")
                        continue;
                    if (!coll.ContainsKey(state)) // test
                    {
                        logger.Warn($"Collection {collKeys[collIndex]} dose not contains state: {state}.");
                        state = "EnableAll";
                        TraversalDevices(state, coll.First(s => s.Key == state));
                        break;
                    }
                    TraversalDevices(state, stateItem);
                }
            }
            currentStates[collIndex] = state;
            System.IO.File.WriteAllText("lastState", JsonConvert.SerializeObject(currentStates));
        }
        static void TraversalDevices(string state, KeyValuePair<string, List<string>> stateItem)
        {
            foreach (var deviceName in stateItem.Value)
            {
                string instanceId = GetDevicesInfo.GetDeviceInstanceId(logger, deviceName, devices);
                if (string.IsNullOrEmpty(instanceId))
                    continue;
                if (stateItem.Key == state)
                {
                    UpdateDeviceStatus("enable", deviceName, instanceId);
                }
                else
                {
                    UpdateDeviceStatus("disable", deviceName, instanceId);
                }
            }
        }

        static void UpdateDeviceStatus(string op, string deviceName, string instanceId)
        {
            if (!string.IsNullOrEmpty(instanceId))
            {
                try
                {
                    var process = new Process
                    {
                        StartInfo = new ProcessStartInfo
                        {
                            FileName = $"{cDir}\\resources\\pnputil.exe",
                            Arguments = $"/{op}-device \"{instanceId}\"",
                            RedirectStandardOutput = true,
                            UseShellExecute = false,
                            CreateNoWindow = true
                        }
                    };
                    process.Start();
                    process.WaitForExit();
                    logger.Info($"[{deviceName}] is {op}d!");
                }
                catch (Exception ex)
                {
                    logger.Error(ex);
                }
            }
        }
    }
}

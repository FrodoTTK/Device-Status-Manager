using DeviceManager;
using DeviceStatusManager.Entitys;
using Newtonsoft.Json;
using NLog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace DeviceStatusManager.Common
{
    public static class ConfigOperator
    {
        public static void ReloadDevicesInfo(Logger logger, string cDir, ref Device[] devices)
        {
            logger.Info("Loading PnP devices information.");
            GetDevicesInfo.RunGetAllPnPDevicesShell(logger, cDir);
            string devicesText = File.ReadAllText(cDir + "\\PnPDevices.json");
            devices = JsonConvert.DeserializeObject<Device[]>(devicesText);
            if (devices.Any())
                logger.Info($"Loaded successfully, total {devices.Count()} entries.");
            else
                logger.Warn($"Loading failed, no entries were loaded.");
        }

        public static void LoadConfig(Logger logger, ref Dictionary<string, Dictionary<string, List<string>>> deviceState)
        {
            try
            {
                var config = File.ReadAllText("config.json");
                deviceState = JsonConvert.DeserializeObject<Dictionary<string, Dictionary<string, List<string>>>>(config);
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
        }
    }
}

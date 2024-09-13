using DeviceStatusManager.Entitys;
using Newtonsoft.Json;
using NLog;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace DeviceStatusManager.Common
{
    public static class GetDevicesInfo
    {
        public static string GetDeviceInstanceId(Logger logger, string deviceDesc,Device[] devices)
        {
            try
            {
                var devicesFounded = devices.Where(d => d.DeviceDescription == deviceDesc && (d.Status == "Started" || d.Status == "Disabled")).ToList();
                if (devicesFounded.Any())
                {
                    logger.Info($"Get [{deviceDesc}] InstanceId successfully!");
                    return devicesFounded[0].InstanceID;
                }
                else
                {
                    logger.Info($"Notfound {deviceDesc} with status of 'Started' or 'Disabled' in 'PnPDevices.json'.");
                    return "";
                }
            }
            catch (System.Exception ex)
            {
                logger.Error(ex);
                throw;
            }
        }
        public static void RunGetAllPnPDevicesShell(Logger logger, string cdir)
        {
            try
            {
                logger.Info("Starting GetAllPnPDevices.ps1.");
                Process proc = new Process();
                proc.StartInfo.FileName = $"{cdir}\\resources\\scripts\\runshell.bat";
                proc.StartInfo.UseShellExecute = false;
                proc.StartInfo.CreateNoWindow = true;
                proc.StartInfo.Verb = "runas";
                proc.Start();
                proc.WaitForExit();
            }
            catch (System.Exception ex)
            {
                logger.Error(ex);
            }
        }
    }
}

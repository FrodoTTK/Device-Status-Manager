using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using IWshRuntimeLibrary;
using NLog;

namespace DeviceStatusManager.Common
{
    public static class CommonFunctions
    {
        public static IWshShortcut StartAtLogin(string startupFolder)
        {
            WshShell shell = new WshShell();
            string shortcutAddress = startupFolder + @"\DeviceStatusManager.lnk";
            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutAddress);
            return shortcut;
        }
        public static void ClearLogFiles(Logger logger)
        {
            string folderPath = Directory.GetCurrentDirectory() + "\\logs";
            DateTime targetDate = DateTime.Now.AddDays(-3);

            try
            {
                // Get all txt files
                string[] files = Directory.GetFiles(folderPath, "*.txt");

                foreach (var file in files)
                {
                    string fileName = Path.GetFileNameWithoutExtension(file);

                    if (DateTime.TryParseExact(fileName, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime fileDate))
                    {
                        if (fileDate < targetDate)
                        {
                            System.IO.File.Delete(file);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
        }
        public class CustomToolStripRenderer : ToolStripProfessionalRenderer
        {
            protected override void OnRenderMenuItemBackground(ToolStripItemRenderEventArgs e)
            {
                if (!e.Item.Enabled)
                {
                    e.Graphics.FillRectangle(new SolidBrush(Color.Transparent), e.Item.ContentRectangle);
                }
                else
                {
                    base.OnRenderMenuItemBackground(e);
                }
            }
        }
    }
}

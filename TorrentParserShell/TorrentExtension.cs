using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SharpShell.Attributes;
using SharpShell.SharpContextMenu;
using SharpShell.SharpInfoTipHandler;

namespace TorrentParser
{
    [ComVisible(true)]
    [COMServerAssociation(AssociationType.ClassOfExtension, ".torrent")]
    public class TorrentExtension : SharpContextMenu
    {
        private static readonly string InfoText;
        private static readonly string LInfoText;
        private static readonly string CopyFail;

        static TorrentExtension()
        {
            

            if (CultureInfo.CurrentCulture.IetfLanguageTag == "zh-CN")
            {
                InfoText = "复制磁力链接";
                LInfoText = "复制磁力链接（带Tracker）";
                CopyFail = "解析磁力链接失败！";
            }
            else if (CultureInfo.CurrentCulture.IetfLanguageTag == "zh-TW" || CultureInfo.CurrentCulture.IetfLanguageTag == "zh-HK")
            {
                InfoText = "複製磁力連結";
                LInfoText = "複製磁力連結（帶Tracker）";
                CopyFail = "解析磁力連結失敗！";
            }
            else
            {
                InfoText = "Copy MagnetURI";
                LInfoText = "Copy MagnetURI (With Tracker)";
                CopyFail = "Torrent parse fail!";
            }
        }

        protected override bool CanShowMenu()
        {
            return true;
        }

        protected override ContextMenuStrip CreateMenu()
        {
            var menu = new ContextMenuStrip();
            var graphics = System.Drawing.Graphics.FromHwnd(IntPtr.Zero);
            var dpi = (int)graphics.DpiX;
            var itemParse = new ToolStripMenuItem
            {
                Name = "tsmP",
                Text = InfoText,
                Image = dpi < 192 ? Properties.Resources.magnet_16 : Properties.Resources.magnet_32
            };
            var itemParseL = new ToolStripMenuItem
            {
                Name = "tsmPL",
                Text = LInfoText,
                Image = dpi < 192 ? Properties.Resources.magnet_16 : Properties.Resources.magnet_32
            };
            itemParse.Click += ItemParse_Click;
            itemParseL.Click += ItemParse_Click;
            menu.Items.Add(itemParse);
            menu.Items.Add(itemParseL);
            return menu;
        }

        private void ItemParse_Click(object sender, EventArgs e)
        {
            var item = sender as ToolStripMenuItem;
            if (item != null)
            {
                try
                {
                    var textBuilder = new StringBuilder();
                    var count = 0;
                    foreach (var filePath in SelectedItemPaths)
                    {
                        Torrent t;
                        if (!Torrent.TryParse(filePath, out t)) continue;
                        textBuilder.AppendLine(item.Name == "tsmPL" ? t.MagnetURI : $"magnet:?xt=urn:btih:{t.SHAHash}");
                        count++;
                    }
                    if (count <= 0)
                    {
                        //MessageBox.Show(CopyFail, "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Helper.ShowBalloon(CopyFail, "Attention", SystemIcons.Error);
                        return;
                    }
                    Clipboard.SetText(textBuilder.ToString().TrimEnd('\n', '\r'));
                }
                catch (Exception ex)
                {
                    //MessageBox.Show($"{CopyFail}:\n{ex.Message}", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Helper.ShowBalloon($"{CopyFail}:\n{ex.Message}", "Attention", SystemIcons.Error);
                }
            }
        }
    }

    [ComVisible(true)]
    [COMServerAssociation(AssociationType.ClassOfExtension, ".torrent")]
    public class TorrentInfoTipHandler : SharpInfoTipHandler
    {
        private static readonly string FolderStr;
        private static readonly string ParseFail;
        private static readonly string Template;
        private static readonly string CannotParse;

        static TorrentInfoTipHandler()
        {
            if (CultureInfo.CurrentCulture.IetfLanguageTag == "zh-CN")
            {
                FolderStr = "文件夹  ";
                CannotParse = "无法解析，文件过大。";
                ParseFail = "种子解析失败！";
                Template = @"种子标题：{0}

文件列表（{1}个文件）：

{2}

创建时间：{3}";
            }
            else if (CultureInfo.CurrentCulture.IetfLanguageTag == "zh-TW" || CultureInfo.CurrentCulture.IetfLanguageTag == "zh-HK")
            {
                FolderStr = "資料夾 ";
                CannotParse = "無法解析，檔案過大。";
                ParseFail = "種子解析失敗！";
                Template = @"種子標題： {0}

檔案清單（{1}個檔案）：

{2}

建立日期： {3}";
            }
            else
            {
                FolderStr = "Folder ";
                CannotParse = "This file is too large!";
                ParseFail = "Torrent parse fail!";
                Template = @"Torrent Title: {0}

File List ({1} Files): 

{2}

Create Time: {3}";
            }
        }

        protected override string GetInfo(RequestedInfoType infoType, bool singleLine)
        {
            switch (infoType)
            {
                case RequestedInfoType.InfoTip:

                    try
                    {
                        var fi = new FileInfo(SelectedItemPath);
                        if (fi.Length > 2000000)
                        {
                            return CannotParse;
                        }
                        Torrent t;
                        using (var ms = new MemoryStream())
                        using (var fs = fi.OpenRead())
                        {
                            fs.CopyTo(ms);
                            if (!Torrent.TryParse(ms.ToArray(), out t))
                                return ParseFail;
                        }
                        return string.Format(Template, t.Name, t.Files.Length, string.Join("\n", t.Files.
                                OrderByDescending(c => c.Length).
                                Take(10).
                                Select(c => $"{c.Name}  <{(c.Length < 1000000 ? c.Length + "Bytes" : c.Length / 1000000f + "MB")}>")), t.CreationDate.ToString("F"));
                    }
                    catch (Exception ex)
                    {
                        return $"{ParseFail}:\n{ex.Message}";
                    }
                case RequestedInfoType.Name:
                    return $"{FolderStr}{Path.GetFileName(SelectedItemPath)}'";
                default:
                    return string.Empty;
            }
        }
    }

    public static class Helper
    {
        public static void ShowBalloon(string body, string title = null, Icon icon = null,int timeout = 3)
        {
            using (var notification = new NotifyIcon
            {
                Visible = true,
                Icon = icon??SystemIcons.Information,
                BalloonTipTitle = title?? "TorrentParser",
                BalloonTipText = body
            })
            {
                notification.ShowBalloonTip(timeout);
            }
        }
    }
}

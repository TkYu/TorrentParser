using System;
using System.Collections.Generic;
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
            if (System.Globalization.CultureInfo.CurrentCulture.IetfLanguageTag.StartsWith("zh", StringComparison.OrdinalIgnoreCase))
            {
                InfoText = "复制磁力链接";
                LInfoText = "复制磁力链接（带Tracker）";
                CopyFail = "解析磁力链接失败！";
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
            var itemParse = new ToolStripMenuItem
            {
                Name = "tsmP",
                Text = InfoText,
                Image = Properties.Resources.Parser
            };
            var itemParseL = new ToolStripMenuItem
            {
                Name = "tsmPL",
                Text = LInfoText,
                Image = Properties.Resources.Parser
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
                        MessageBox.Show(CopyFail, "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    Clipboard.SetText(textBuilder.ToString().TrimEnd('\n', '\r'));
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"{CopyFail}:\n{ex.Message}", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            if (System.Globalization.CultureInfo.CurrentCulture.IetfLanguageTag.StartsWith("zh", StringComparison.OrdinalIgnoreCase))
            {
                FolderStr = "文件夹  ";
                CannotParse = "无法解析，文件过大。";
                ParseFail = "种子解析失败！";
                Template = @"种子标题：{0}

文件列表（{1}个文件）：

{2}

创建时间：{3}";
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
}

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
        private static readonly string CopyFail;
        private static readonly string CopySuccess;

        static TorrentExtension()
        {
            if (System.Globalization.CultureInfo.CurrentCulture.IetfLanguageTag.StartsWith("zh", StringComparison.OrdinalIgnoreCase))
            {
                InfoText = "复制磁力链接...";
                CopyFail = "解析磁力链接失败！";
                CopySuccess = "成功复制到剪贴板！";
            }
            else
            {
                InfoText = "Copy MagnetURI...";
                CopyFail = "Torrent parse fail!";
                CopySuccess = "Copy Successfully!";
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
                Text = InfoText,
                Image = Properties.Resources.Parser
            };
            itemParse.Click += ItemParse_Click;
            menu.Items.Add(itemParse);
            return menu;
        }

        private void ItemParse_Click(object sender, EventArgs e)
        {
            try
            {
                var messageBuilder = new StringBuilder();
                var textBuilder = new StringBuilder();
                int count = 0;
                foreach (var filePath in SelectedItemPaths)
                {
                    Torrent t;
                    if (!Torrent.TryParse(filePath, out t)) continue;
                    messageBuilder.AppendLine($"{Path.GetFileNameWithoutExtension(filePath)} = {t.SHAHash}");
                    textBuilder.AppendLine(t.MagnetURI);
                    count++;
                }
                if (count<=0)
                {
                    MessageBox.Show(CopyFail, "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                Clipboard.SetText(textBuilder.ToString().TrimEnd('\n','\r'));
                MessageBox.Show($"{messageBuilder}{CopySuccess}", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{CopyFail}:\n{ex.Message}", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                        var bt = File.ReadAllBytes(SelectedItemPath);
                        if (bt.Length > 2000000)
                        {
                            return CannotParse;
                        }
                        var t = new Torrent(bt);
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

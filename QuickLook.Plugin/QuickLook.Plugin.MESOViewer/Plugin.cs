using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using QuickLook.Common.Helpers;
using QuickLook.Common.Plugin;
using QuickLook.Plugin.ImageViewer;

namespace QuickLook.Plugin.MESOViewer
{
    public class Plugin : IViewer
    {
        [DllImport(@"MESOImageGenerator.dll", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern void mesotoimage(IntPtr meso_path, IntPtr img_out_path);

        private ImagePanel _panel;
        private MetaProvider _meta;
        private Uri _mesoImageURI;

        public int Priority => 0;

        public void Init()
        {
        }

        public bool CanHandle(string path)
        {
            return !Directory.Exists(path) && path.ToLower().EndsWith(".meso");
        }

        public void Prepare(string path, ContextObject context)
        {
            var imgPath = GetTempPNG();

            mesotoimage(Marshal.StringToHGlobalAnsi(path), Marshal.StringToHGlobalAnsi(imgPath));

            _mesoImageURI = FilePathToFileUrl(imgPath);;

            _meta = new MetaProvider(imgPath);
            var size = _meta.GetSize();
            if (!size.IsEmpty)
                context.SetPreferredSizeFit(size, 0.8);
            else
                context.PreferredSize = new Size(800, 600);

            context.Theme = (Themes) SettingHelper.Get("LastTheme", 1, "QuickLook.Plugin.ImageViewer");
        }

        public void View(string path, ContextObject context)
        {
            _panel = new ImagePanel
            {
                Meta = _meta,
                ContextObject = context
            };

            var size = _meta.GetSize();

            context.ViewerContent = _panel;
            context.Title = size.IsEmpty
                ? $"{Path.GetFileName(path)}"
                : $"{size.Width}×{size.Height}: {Path.GetFileName(path)}";

            _panel.ImageUriSource = _mesoImageURI;
        }

        public void Cleanup()
        {
            File.Delete(_mesoImageURI.AbsolutePath);
            _mesoImageURI = null;
            _panel?.Dispose();
            _panel = null;
        }
        private string GetTempPNG()
        {
            var origPath = Path.GetTempFileName();
            File.Delete(origPath);
            return origPath.Replace(".tmp", ".png");
        }

        public static Uri FilePathToFileUrl(string filePath)
        {
            var uri = new StringBuilder();
            foreach (var v in filePath)
                if (v >= 'a' && v <= 'z' || v >= 'A' && v <= 'Z' || v >= '0' && v <= '9' ||
                    v == '+' || v == '/' || v == ':' || v == '.' || v == '-' || v == '_' || v == '~' ||
                    v > '\x80')
                    uri.Append(v);
                else if (v == Path.DirectorySeparatorChar || v == Path.AltDirectorySeparatorChar)
                    uri.Append('/');
                else
                    uri.Append($"%{(int)v:X2}");
            if (uri.Length >= 2 && uri[0] == '/' && uri[1] == '/') // UNC path
                uri.Insert(0, "file:");
            else
                uri.Insert(0, "file:///");

            try
            {
                return new Uri(uri.ToString());
            }
            catch
            {
                return null;
            }
        }
    }
}
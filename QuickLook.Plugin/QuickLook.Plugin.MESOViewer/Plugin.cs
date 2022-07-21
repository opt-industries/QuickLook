using ICSharpCode.AvalonEdit;
using ICSharpCode.AvalonEdit.Highlighting;
using ICSharpCode.AvalonEdit.Highlighting.Xshd;
using QuickLook.Common.Helpers;
using QuickLook.Common.Plugin;
using QuickLook.Plugin.ImageViewer;
using QuickLook.Plugin.TextViewer;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Xml;

namespace QuickLook.Plugin.MESOViewer
{
    public class Plugin : IViewer
    {
        [DllImport(@"MESOImageGenerator.dll", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr mesotoimage(IntPtr meso_path, IntPtr img_out_path);

        private ImagePanel _imageViewerPanel;
        private TextViewerPanel _textViewerPanel;
        private MetaProvider _meta;
        private Uri _mesoPreviewURI;
        private string _extension;
        private static HighlightingManager _hlmLight;
        private static HighlightingManager _hlmDark;

        public int Priority => 0;

        public void Init()
        {
            var _ = new TextEditor();

            _hlmLight = GetHighlightingManager("Light");
            _hlmDark = GetHighlightingManager("Dark");
        }

        public bool CanHandle(string path)
        {
            return !Directory.Exists(path) && path.ToLower().EndsWith(".meso");
        }

        public void Prepare(string path, ContextObject context)
        {
            _mesoPreviewURI = FilePathToFileUrl(
                Marshal.PtrToStringAnsi(
                    mesotoimage(
                        Marshal.StringToHGlobalAnsi(path),
                        Marshal.StringToHGlobalAnsi(GetTempPath())
                    )
                )
            );
            _extension = Path.GetExtension(_mesoPreviewURI.ToString());

            if (_extension == ".json")
            {
                context.PreferredSize = new Size(800, 600);  // text preview
            }
            else  // image preview
            {
                _meta = new MetaProvider(_mesoPreviewURI.AbsolutePath);
                var size = _meta.GetSize();
                if (!size.IsEmpty)
                    context.SetPreferredSizeFit(size, 0.8);
                else
                    context.PreferredSize = new Size(800, 600);

                context.Theme = (Themes)SettingHelper.Get("LastTheme", 1, "QuickLook.Plugin.ImageViewer");
            }
        }

        private void ViewImage(string path, ContextObject context)
        {
            _imageViewerPanel = new ImagePanel
            {
                Meta = _meta,
                ContextObject = context
            };

            var size = _meta.GetSize();

            context.ViewerContent = _imageViewerPanel;
            context.Title = size.IsEmpty
                ? $"{Path.GetFileName(path)}"
                : $"{size.Width}×{size.Height}: {Path.GetFileName(path)}";

            _imageViewerPanel.ImageUriSource = _mesoPreviewURI;
        }

        private void ViewText(ContextObject context)
        {
            _textViewerPanel = new TextViewerPanel(_mesoPreviewURI.AbsolutePath, context);
            AssignHighlightingManager(_textViewerPanel, context);
            context.ViewerContent = _textViewerPanel;
        }

        public void View(string path, ContextObject context)
        {
            if (_extension == ".json")
            {
                ViewText(context);
            }
            else
            {
                ViewImage(path, context);
            }
        }

        public void Cleanup()
        {
            File.Delete(_mesoPreviewURI.AbsolutePath);
            _mesoPreviewURI = null;
            _extension = null;
            _meta = null;
            _imageViewerPanel?.Dispose();
            _imageViewerPanel = null;
            _textViewerPanel?.Dispose();
            _textViewerPanel = null;
        }
        private string GetTempPath()
        {
            var tmpPath = Path.GetTempFileName();
            File.Delete(tmpPath);
            return tmpPath;
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
        private HighlightingManager GetHighlightingManager(string dirName)
        {
            var hlm = new HighlightingManager();

            var assemblyPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            if (string.IsNullOrEmpty(assemblyPath))
                return hlm;

            var syntaxPath = Path.Combine(assemblyPath, "Syntax", dirName);
            if (!Directory.Exists(syntaxPath))
                return hlm;

            foreach (var file in Directory.EnumerateFiles(syntaxPath, "*.xshd"))
            {
                Debug.WriteLine(file);
                var ext = Path.GetFileNameWithoutExtension(file);
                using (Stream s = File.OpenRead(Path.GetFullPath(file)))
                using (var reader = new XmlTextReader(s))
                {
                    var xshd = HighlightingLoader.LoadXshd(reader);
                    var highlightingDefinition = HighlightingLoader.Load(xshd, hlm);
                    if (xshd.Extensions.Count > 0)
                        hlm.RegisterHighlighting(ext, xshd.Extensions.ToArray(), highlightingDefinition);
                }
            }

            return hlm;
        }
        private void AssignHighlightingManager(TextViewerPanel tvp, ContextObject context)
        {
            var isDark = (context.Theme == Themes.Dark) | OSThemeHelper.AppsUseDarkTheme() | false;
            tvp.HighlightingManager = isDark ? _hlmDark : _hlmLight;
        }
    }
}
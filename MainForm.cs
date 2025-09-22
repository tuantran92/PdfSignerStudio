using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace PdfSignerStudio
{
    public partial class MainForm : Form
    {
        ProjectState state = new();
        WebView2 web = new() { Dock = DockStyle.Fill };

        Panel topBar = new() { Height = 40, Dock = DockStyle.Top };
        ComboBox pageBox = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 80 };
        Button btnOpenDocx = new() { Text = "Open" };
        Button btnSaveJson = new() { Text = "Save JSON" };
        Button btnLoadJson = new() { Text = "Load JSON" };
        Button btnExport = new() { Text = "Export PDF" };
        Button btnZoomOut = new() { Text = "−" };
        Button btnZoomIn = new() { Text = "+" };
        Button btnTplFolder = new() { Text = "Templates…" };
        Label info = new() { AutoSize = true, ForeColor = Color.DimGray };
        ToolTip toolTip = new(); // Tooltip component

        // --- STATUS STRIP CONTROLS ---
        StatusStrip statusBar = new();
        ToolStripStatusLabel lblStatus = new() { Spring = true, TextAlign = ContentAlignment.MiddleLeft };
        ToolStripStatusLabel lblFileName = new() { BorderSides = ToolStripStatusLabelBorderSides.Left, BorderStyle = Border3DStyle.Etched, Padding = new Padding(5, 0, 5, 0) };
        ToolStripStatusLabel lblFieldCount = new() { BorderSides = ToolStripStatusLabelBorderSides.Left, BorderStyle = Border3DStyle.Etched, Padding = new Padding(5, 0, 5, 0) };
        ToolStripStatusLabel lblCoords = new() { BorderSides = ToolStripStatusLabelBorderSides.Left, BorderStyle = Border3DStyle.Etched, Padding = new Padding(5, 0, 5, 0) };


        // ===== Template library =====
        List<TemplateDef> templates = new();
        string templatesDir = Path.Combine(AppContext.BaseDirectory, "Template");
        FileSystemWatcher? tplWatcher;

        readonly FlowLayoutPanel toolbar = new()
        {
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            WrapContents = false,
            Padding = new Padding(0),
            Margin = new Padding(0)
        };

        void CenterToolbar()
        {
            if (toolbar.Parent == null) return;
            var pref = toolbar.PreferredSize;
            toolbar.Location = new Point(
                Math.Max(0, (topBar.ClientSize.Width - pref.Width) / 2),
                Math.Max(0, (topBar.ClientSize.Height - pref.Height) / 2)
            );
        }

        void PositionInfo()
        {
            if (info.Parent == null) return;
            info.Left = topBar.ClientSize.Width - info.Width - 10;
            info.Top = (topBar.ClientSize.Height - info.Height) / 2;
        }


        public MainForm()
        {
            InitializeComponent();
            SetupUi();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            // Đây là nơi an toàn để cập nhật giao diện lần đầu
            UpdateStatus("Ready. Please open a DOCX or PDF file.");
            UpdateFieldCount();
            UpdateFileName("No file open");
            UpdateCoordinates(0, 0);
        }

        private void SetupUi()
        {
            Text = "PdfSignerStudio (Word Interop + WebView2 + iText7)";
            ClientSize = new Size(1280, 820);
            StartPosition = FormStartPosition.CenterScreen;

            topBar.Height = 44;
            topBar.Dock = DockStyle.Top;

            btnZoomOut.Width = 30; btnZoomIn.Width = 30;
            info.AutoSize = true;
            info.Anchor = AnchorStyles.Top | AnchorStyles.Right;

            toolbar.Controls.AddRange(new Control[]
            {
                btnOpenDocx, pageBox, btnZoomOut, btnZoomIn,
                btnSaveJson, btnLoadJson, btnExport, btnTplFolder
            });

            topBar.Controls.Add(toolbar);
            topBar.Controls.Add(info);

            Controls.Add(web);
            Controls.Add(topBar);

            // Events
            btnOpenDocx.Click += OnOpenFile;
            pageBox.SelectedIndexChanged += (_, __) => SyncPageToWeb();
            btnSaveJson.Click += (_, __) => SaveJson();
            btnLoadJson.Click += (_, __) => LoadJson();
            btnExport.Click += (_, __) => ExportPdf();
            btnZoomIn.Click += async (_, __) => await web.CoreWebView2?.ExecuteScriptAsync("zoomIn()")!;
            btnZoomOut.Click += async (_, __) => await web.CoreWebView2?.ExecuteScriptAsync("zoomOut()")!;
            btnTplFolder.Click += (_, __) =>
            {
                Directory.CreateDirectory(templatesDir);
                System.Diagnostics.Process.Start("explorer.exe", templatesDir);
            };

            // Gán ToolTips cho các nút
            toolTip.SetToolTip(btnOpenDocx, "Open a Word (.docx) or PDF file");
            toolTip.SetToolTip(btnSaveJson, "Save current field layout to a .json file");
            toolTip.SetToolTip(btnLoadJson, "Load a field layout from a .json file");
            toolTip.SetToolTip(btnExport, "Export the final PDF with signature fields");
            toolTip.SetToolTip(btnZoomIn, "Zoom In");
            toolTip.SetToolTip(btnZoomOut, "Zoom Out");
            toolTip.SetToolTip(btnTplFolder, "Open the templates folder");
            toolTip.SetToolTip(pageBox, "Select a page to view");

            Load += MainForm_Load; // Đăng ký sự kiện Load
            Load += (_, __) => { CenterToolbar(); PositionInfo(); };
            Resize += (_, __) => { CenterToolbar(); PositionInfo(); };
            topBar.Resize += (_, __) => { CenterToolbar(); PositionInfo(); };
            info.SizeChanged += (_, __) => PositionInfo();

            // Thêm StatusStrip vào Form
            statusBar.Items.AddRange(new ToolStripItem[] { lblStatus, lblFileName, lblFieldCount, lblCoords });
            Controls.Add(statusBar);
        }

        // --- Các hàm cập nhật StatusStrip ---
        private void UpdateStatus(string message)
        {
            if (IsHandleCreated)
            {
                BeginInvoke(new Action(() => lblStatus.Text = message));
            }
        }

        private void UpdateFieldCount()
        {
            if (IsHandleCreated)
            {
                BeginInvoke(new Action(() => lblFieldCount.Text = $"{state.Fields.Count} fields"));
            }
        }

        private void UpdateFileName(string filePath)
        {
            if (IsHandleCreated)
            {
                BeginInvoke(new Action(() => lblFileName.Text = Path.GetFileName(filePath)));
            }
        }

        private void UpdateCoordinates(float x, float y)
        {
            if (IsHandleCreated)
            {
                BeginInvoke(new Action(() => lblCoords.Text = $"X: {x:F1}, Y: {y:F1} pt"));
            }
        }

        async Task PushAllFieldsToJs()
        {
            if (web.CoreWebView2 == null) return;
            var allFields = state.Fields
                .OrderBy(f => f.Page)
                .ThenBy(f => f.Name)
                .Select(f => new { id = f.Id, name = f.Name, page = f.Page });

            string json = JsonSerializer.Serialize(allFields);
            await web.CoreWebView2.ExecuteScriptAsync($"setAddedFields({json});");
        }

        static Task<T> RunSTA<T>(Func<T> func)
        {
            var tcs = new TaskCompletionSource<T>();
            var th = new Thread(() =>
            {
                try { tcs.SetResult(func()); }
                catch (Exception ex) { tcs.SetException(ex); }
            });
            th.SetApartmentState(ApartmentState.STA);
            th.Start();
            return tcs.Task;
        }

        async void OnOpenFile(object? sender, EventArgs e)
        {
            using var ofd = new OpenFileDialog
            {
                Filter = "Word/PDF (*.docx;*.pdf)|*.docx;*.pdf|All files (*.*)|*.*",
                FilterIndex = 1,
                Title = "Open Word/PDF",
                CheckFileExists = true,
                RestoreDirectory = true
            };
            if (ofd.ShowDialog() != DialogResult.OK) return;

            string ext = Path.GetExtension(ofd.FileName).ToLowerInvariant();

            state = new ProjectState();
            _ = PushAllFieldsToJs();
            string outDir = Path.Combine(Path.GetTempPath(), "PdfSignerStudio");
            Directory.CreateDirectory(outDir);

            try
            {
                if (ext == ".docx")
                {
                    SplashForm? splash = null;
                    try
                    {
                        splash = new SplashForm();
                        splash.Show(this);
                        Application.DoEvents();

                        UpdateStatus("Converting DOCX → PDF with Microsoft Word...");
                        state.SourceDocx = ofd.FileName;

                        state.TempPdf = await RunSTA(() =>
                            PdfService.ConvertDocxToPdfWithWord(ofd.FileName, outDir));
                    }
                    finally
                    {
                        splash?.Close();
                        splash?.Dispose();
                    }
                }
                else
                {
                    UpdateStatus("Loading PDF...");
                    state.SourceDocx = null;
                    string dest = Path.Combine(outDir, Path.GetFileName(ofd.FileName));
                    try
                    {
                        File.Copy(ofd.FileName, dest, overwrite: true);
                        state.TempPdf = dest;
                    }
                    catch
                    {
                        state.TempPdf = ofd.FileName;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Open failed: " + ex.Message);
                UpdateStatus("Open failed.");
                return;
            }

            UpdateStatus("Loading preview...");
            await EnsureWebReady();

            web.CoreWebView2.WebMessageReceived -= WebMessageReceived;
            web.CoreWebView2.WebMessageReceived += WebMessageReceived;

            var host = "files.local";
            var pdfFolder = Path.GetDirectoryName(state.TempPdf!)!;
            pdfFolder = Path.GetFullPath(pdfFolder);

            if (!Directory.Exists(pdfFolder))
                throw new DirectoryNotFoundException(pdfFolder);

            var cwv2 = web.CoreWebView2;
            try { cwv2.ClearVirtualHostNameToFolderMapping(host); } catch { }
            cwv2.SetVirtualHostNameToFolderMapping(
                host, pdfFolder, CoreWebView2HostResourceAccessKind.Allow);

            var pdfUri = $"https://{host}/{Path.GetFileName(state.TempPdf!)}";
            var html = BuildPdfHtml(pdfUri);

            web.CoreWebView2.NavigationCompleted -= OnWebReady;
            web.CoreWebView2.NavigationCompleted += OnWebReady;

            web.CoreWebView2.NavigateToString(html);

            UpdateFileName(ofd.FileName);
            UpdateStatus("Ready. Drag, drop, nudge, snap, rename inline, flip pages with mouse/PageUp-Down.");
            info.Text = "";
        }

        private async void OnWebReady(object? sender, CoreWebView2NavigationCompletedEventArgs e)
        {
            web.CoreWebView2.NavigationCompleted -= OnWebReady;

            LoadTemplates();
            await PushTemplatesToJs();
            await PushAllFieldsToJs();
            SetupTplWatcher();
            UpdateFieldCount();

            if (pageBox.SelectedIndex >= 0)
            {
                int page = pageBox.SelectedIndex + 1;
                PushFieldsToJs(page);
            }
        }

        void SetupTplWatcher()
        {
            try
            {
                tplWatcher?.Dispose();
                Directory.CreateDirectory(templatesDir);

                tplWatcher = new FileSystemWatcher(templatesDir, "*.json")
                {
                    IncludeSubdirectories = false,
                    NotifyFilter = NotifyFilters.FileName |
                                   NotifyFilters.LastWrite |
                                   NotifyFilters.Size
                };

                tplWatcher.Changed += OnTplChanged;
                tplWatcher.Created += OnTplChanged;
                tplWatcher.Deleted += OnTplChanged;
                tplWatcher.Renamed += OnTplRenamed;
                tplWatcher.EnableRaisingEvents = true;
            }
            catch { }
        }

        async void OnTplChanged(object? sender, FileSystemEventArgs e)
        {
            try
            {
                await Task.Delay(250);
                if (IsDisposed) return;
                BeginInvoke(new Action(async () =>
                {
                    LoadTemplates();
                    await PushTemplatesToJs();
                    UpdateStatus("Templates reloaded.");
                }));
            }
            catch { }
        }
        void OnTplRenamed(object? s, RenamedEventArgs e) => OnTplChanged(s, e);

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            tplWatcher?.Dispose();
            base.OnFormClosed(e);
        }

        private void WebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                var json = e.TryGetWebMessageAsString();
                if (string.IsNullOrWhiteSpace(json)) return;

                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;
                var type = root.GetProperty("type").GetString();

                switch (type)
                {
                    case "meta":
                        {
                            int num = root.GetProperty("numPages").GetInt32();
                            int page = root.GetProperty("page").GetInt32();
                            if (pageBox.Items.Count != num)
                            {
                                pageBox.Items.Clear();
                                for (int i = 1; i <= num; i++) pageBox.Items.Add($"Page {i}");
                                pageBox.SelectedIndex = Math.Max(0, page - 1);
                            }
                            PushFieldsToJs(page);
                            break;
                        }
                    case "addField":
                        {
                            int page = root.GetProperty("page").GetInt32();
                            var r = root.GetProperty("rect");
                            float x = r.GetProperty("x").GetSingle();
                            float y = r.GetProperty("y").GetSingle();
                            float w = r.GetProperty("w").GetSingle();
                            float h = r.GetProperty("h").GetSingle();
                            bool req = r.TryGetProperty("required", out var rq) ? rq.GetBoolean() : true;
                            string name = root.TryGetProperty("name", out var nn) ? (nn.GetString() ?? "") : "";

                            if (string.IsNullOrWhiteSpace(name))
                                name = $"Signature_{state.Fields.Count(f => f.Type == "signature") + 1}";

                            state.Fields.Add(new FormFieldDef(name, "signature", page, new RectFpt(x, y, w, h), req));
                            UpdateStatus($"Added {name} on page {page}");
                            PushFieldsToJs(page);
                            UpdateFieldCount();
                            _ = PushAllFieldsToJs();
                            break;
                        }
                    case "updateField":
                        {
                            string id = root.GetProperty("id").GetString()!;
                            int page = root.GetProperty("page").GetInt32();
                            var r = root.GetProperty("rect");
                            float x = r.GetProperty("x").GetSingle();
                            float y = r.GetProperty("y").GetSingle();
                            float w = r.GetProperty("w").GetSingle();
                            float h = r.GetProperty("h").GetSingle();

                            var f = state.Fields.FirstOrDefault(t => t.Id == id);
                            if (f != null)
                            {
                                state.Fields[state.Fields.IndexOf(f)] = f with { Rect = new RectFpt(x, y, w, h), Page = page };
                                PushFieldsToJs(page);
                                _ = PushAllFieldsToJs();
                            }
                            break;
                        }
                    case "deleteField":
                        {
                            string id = root.GetProperty("id").GetString()!;
                            int page = root.GetProperty("page").GetInt32();
                            state.Fields.RemoveAll(f => f.Id == id);
                            UpdateStatus("Field deleted.");
                            PushFieldsToJs(page);
                            UpdateFieldCount();
                            _ = PushAllFieldsToJs();
                            break;
                        }
                    case "renameField":
                        {
                            string id = root.GetProperty("id").GetString()!;
                            string newName = root.GetProperty("name").GetString() ?? "";
                            int page = root.GetProperty("page").GetInt32();

                            if (string.IsNullOrWhiteSpace(newName)) break;
                            string baseName = newName.Trim();
                            string name = baseName; int idx = 1;
                            while (state.Fields.Any(f => f.Name.Equals(name, StringComparison.OrdinalIgnoreCase) && f.Id != id))
                                name = $"{baseName}_{idx++}";

                            var f = state.Fields.FirstOrDefault(t => t.Id == id);
                            if (f != null)
                            {
                                state.Fields[state.Fields.IndexOf(f)] = f with { Name = name };
                                PushFieldsToJs(page);
                                _ = PushAllFieldsToJs();
                            }
                            break;
                        }
                    case "toggleRequired":
                        {
                            string id = root.GetProperty("id").GetString()!;
                            int page = root.GetProperty("page").GetInt32();
                            var f = state.Fields.FirstOrDefault(t => t.Id == id);
                            if (f != null)
                            {
                                state.Fields[state.Fields.IndexOf(f)] = f with { Required = !f.Required };
                                PushFieldsToJs(page);
                            }
                            break;
                        }
                    case "mouseMove":
                        {
                            var p = root.GetProperty("pt");
                            float x = p.GetProperty("x").GetSingle();
                            float y = p.GetProperty("y").GetSingle();
                            UpdateCoordinates(x, y);
                            break;
                        }

                    case "saveTemplate":
                        {
                            var t = root.GetProperty("template");
                            string name = t.GetProperty("name").GetString() ?? "Unnamed";
                            if (string.IsNullOrWhiteSpace(name)) break;

                            string Safe(string s)
                            {
                                foreach (var c in Path.GetInvalidFileNameChars())
                                    s = s.Replace(c, '_');
                                return s.Trim();
                            }

                            var items = new List<TemplateField>();
                            foreach (var it in t.GetProperty("items").EnumerateArray())
                            {
                                string iname = it.GetProperty("name").GetString() ?? "Field";
                                float w = it.GetProperty("w").GetSingle();
                                float h = it.GetProperty("h").GetSingle();
                                bool req = it.TryGetProperty("required", out var rq) ? rq.GetBoolean() : true;
                                float dx = it.TryGetProperty("dx", out var dxv) ? dxv.GetSingle() : 0f;
                                float dy = it.TryGetProperty("dy", out var dyv) ? dyv.GetSingle() : 0f;
                                items.Add(new TemplateField(iname, w, h, req, dx, dy));
                            }

                            Directory.CreateDirectory(templatesDir);
                            var def = new TemplateDef(name, items);
                            var tplJson = JsonSerializer.Serialize(def, new JsonSerializerOptions { WriteIndented = true });
                            var path = Path.Combine(templatesDir, Safe(name) + ".json");
                            File.WriteAllText(path, tplJson);

                            LoadTemplates();
                            _ = PushTemplatesToJs();
                            UpdateStatus($"Saved template: {name}");
                            break;
                        }

                    case "deleteTemplate":
                        {
                            string name = root.GetProperty("name").GetString() ?? "";
                            if (string.IsNullOrWhiteSpace(name)) break;

                            string Safe(string s)
                            {
                                foreach (var c in Path.GetInvalidFileNameChars())
                                    s = s.Replace(c, '_');
                                return s.Trim();
                            }
                            var path = Path.Combine(templatesDir, Safe(name) + ".json");
                            if (File.Exists(path)) File.Delete(path);

                            LoadTemplates();
                            _ = PushTemplatesToJs();
                            UpdateStatus($"Deleted template: {name}");
                            break;
                        }
                }
            }
            catch { }
        }

        private async void SyncPageToWeb()
        {
            if (web.CoreWebView2 == null || pageBox.SelectedIndex < 0) return;
            int page = pageBox.SelectedIndex + 1;
            await web.CoreWebView2.ExecuteScriptAsync($"setPage({page});");
            PushFieldsToJs(page);
        }

        private async void PushFieldsToJs(int page)
        {
            if (web.CoreWebView2 == null) return;

            for (int i = 0; i < state.Fields.Count; i++)
            {
                if (string.IsNullOrEmpty(state.Fields[i].Id))
                    state.Fields[i] = state.Fields[i] with { };
            }

            var list = state.Fields.Where(f => f.Page == page).Select(f => new
            {
                id = f.Id,
                name = f.Name,
                x = f.Rect.X,
                y = f.Rect.Y,
                w = f.Rect.W,
                h = f.Rect.H,
                required = f.Required
            });
            string json = JsonSerializer.Serialize(list);
            await web.CoreWebView2.ExecuteScriptAsync($"setFields({json});");
        }

        void LoadTemplates()
        {
            templates.Clear();
            Directory.CreateDirectory(templatesDir);

            if (!Directory.EnumerateFiles(templatesDir, "*.json").Any())
            {
                var demo1 = new TemplateDef("Signature 120×60", new List<TemplateField>
                {
                    new TemplateField("Signature", 120, 60, true, 0, 0)
                });
                var demo2 = new TemplateDef("Director + Accountant", new List<TemplateField>
                {
                    new TemplateField("Director", 140, 70, true, 0, 0),
                    new TemplateField("Accountant", 140, 70, true, 160, 0),
                });
                File.WriteAllText(Path.Combine(templatesDir, "Signature_120x60.json"),
                    JsonSerializer.Serialize(demo1, new JsonSerializerOptions { WriteIndented = true }));
                File.WriteAllText(Path.Combine(templatesDir, "Director_Accountant.json"),
                    JsonSerializer.Serialize(demo2, new JsonSerializerOptions { WriteIndented = true }));
            }

            foreach (var f in Directory.EnumerateFiles(templatesDir, "*.json"))
            {
                try
                {
                    var t = JsonSerializer.Deserialize<TemplateDef>(File.ReadAllText(f));
                    if (t?.Items != null && t.Items.Count > 0)
                        templates.Add(t);
                }
                catch { }
            }
        }

        async Task PushTemplatesToJs()
        {
            if (web.CoreWebView2 == null) return;
            var payload = templates.Select(t => new
            {
                name = t.Name,
                items = t.Items.Select(i => new { name = i.Name, w = i.W, h = i.H, required = i.Required, dx = i.Dx, dy = i.Dy })
            });
            string json = JsonSerializer.Serialize(payload);
            await web.CoreWebView2.ExecuteScriptAsync($"setTemplates({json});");
        }

        void SaveJson()
        {
            using var sfd = new SaveFileDialog { Filter = "JSON (*.json)|*.json" };
            if (sfd.ShowDialog() != DialogResult.OK) return;
            File.WriteAllText(sfd.FileName, JsonSerializer.Serialize(state, new JsonSerializerOptions { WriteIndented = true }));
            UpdateStatus("Saved JSON project.");
        }

        async void LoadJson()
        {
            using var ofd = new OpenFileDialog
            {
                Filter = "JSON (*.json)|*.json|All files (*.*)|*.*",
                FilterIndex = 1,
                Title = "Load JSON",
                CheckFileExists = true,
                RestoreDirectory = true
            };
            if (ofd.ShowDialog() != DialogResult.OK) return;

            try
            {
                state = JsonSerializer.Deserialize<ProjectState>(File.ReadAllText(ofd.FileName)) ?? new ProjectState();

                for (int i = 0; i < state.Fields.Count; i++)
                    if (string.IsNullOrEmpty(state.Fields[i].Id))
                        state.Fields[i] = state.Fields[i] with { };

                UpdateFileName(ofd.FileName);
                UpdateStatus("Loaded JSON project successfully.");
                info.Text = "";

                if (!string.IsNullOrEmpty(state.TempPdf) && File.Exists(state.TempPdf))
                {
                    await EnsureWebReady();

                    web.CoreWebView2.WebMessageReceived -= WebMessageReceived;
                    web.CoreWebView2.WebMessageReceived += WebMessageReceived;

                    var host = "files.local";
                    var pdfFolder = Path.GetDirectoryName(state.TempPdf!)!;
                    pdfFolder = Path.GetFullPath(pdfFolder);

                    if (!Directory.Exists(pdfFolder))
                        throw new DirectoryNotFoundException(pdfFolder);

                    var cwv2 = web.CoreWebView2;
                    try { cwv2.ClearVirtualHostNameToFolderMapping(host); } catch { }
                    cwv2.SetVirtualHostNameToFolderMapping(
                        host, pdfFolder, CoreWebView2HostResourceAccessKind.Allow);

                    var pdfUri = $"https://{host}/{Path.GetFileName(state.TempPdf!)}";
                    var html = BuildPdfHtml(pdfUri);

                    web.CoreWebView2.NavigationCompleted -= OnWebReady;
                    web.CoreWebView2.NavigationCompleted += OnWebReady;

                    web.CoreWebView2.NavigateToString(html);
                }
                else
                {
                    await PushAllFieldsToJs();
                    MessageBox.Show(
                        "JSON file does not contain a valid path to a PDF (TempPdf). Please open a DOCX/PDF first, or edit the 'TempPdf' path in the JSON.",
                        "Missing PDF", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Load JSON failed: " + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        void ExportPdf()
        {
            if (state.TempPdf == null)
            {
                MessageBox.Show("No PDF file is currently open.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using var sfd = new SaveFileDialog { Filter = "PDF (*.pdf)|*.pdf", Title = "Save PDF File" };
            if (sfd.ShowDialog() != DialogResult.OK) return;

            try
            {
                PdfService.AddSignatureFields(state.TempPdf, sfd.FileName, state.Fields);
                UpdateStatus($"Exported: {sfd.FileName}");
                MessageBox.Show($"Successfully exported PDF!\nSaved at: {sfd.FileName}", "Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (IOException ioEx)
            {
                MessageBox.Show(
                    $"Cannot save the file.\n\nThe file '{Path.GetFileName(sfd.FileName)}' might be open in another program (like Adobe Reader, Chrome, etc.).\n\nPlease close that file and try again.",
                    "File Write Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
                UpdateStatus("Export failed: File in use.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred during export:\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateStatus("Export failed.");
            }
        }

        private async Task EnsureWebReady()
        {
            if (web.CoreWebView2 == null)
                await web.EnsureCoreWebView2Async();
        }
        private static string HtmlFilePath()
        {
            return Path.Combine(AppContext.BaseDirectory, "Web", "index.html");
        }

        private string BuildPdfHtml(string pdfFileUri)
        {
            var path = HtmlFilePath();
            if (!File.Exists(path))
                throw new FileNotFoundException("Web\\index.html not found. Please create the file and set 'Copy to Output' = 'Copy if newer'.", path);

            var tpl = File.ReadAllText(path);
            return tpl.Replace("__PDF_URL__", pdfFileUri);
        }
    }
}
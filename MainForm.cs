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
        #region State and Core Components

        private ProjectState state = new();
        private readonly WebView2 web = new() { Dock = DockStyle.Fill };
        private List<TemplateDef> templates = new();
        private readonly string templatesDir = Path.Combine(AppContext.BaseDirectory, "Template");
        private FileSystemWatcher? tplWatcher;

        // === Undo/Redo ===
        private readonly Stack<ProjectState> _undo = new();
        private readonly Stack<ProjectState> _redo = new();
        private int _currentPage = 1;

        #endregion

        #region UI Controls

        // === KHAI BÁO LẠI UI THEO CÁCH ĐƠN GIẢN NHẤT ===
        private ToolStrip topToolstrip;
        private Panel toolHost;
        private ToolStripButton btnOpen, btnSaveJson, btnLoadJson, btnExport;
        private ToolStripButton btnZoomIn, btnZoomOut;
        private ToolStripButton btnTplFolder;
        private ToolStripButton btnUndo, btnRedo;
        private bool _isDirty = false;
        private ToolStripButton btnGrid;
        private StatusStrip statusBar = new();
        private ToolStripStatusLabel lblStatus = new() { Spring = true, TextAlign = ContentAlignment.MiddleLeft };
        private ToolStripStatusLabel lblFileName = new() { BorderSides = ToolStripStatusLabelBorderSides.Left, BorderStyle = Border3DStyle.Etched, Padding = new Padding(5, 0, 5, 0) };
        private ToolStripStatusLabel lblFieldCount = new() { BorderSides = ToolStripStatusLabelBorderSides.Left, BorderStyle = Border3DStyle.Etched, Padding = new Padding(5, 0, 5, 0) };
        private ToolStripStatusLabel lblCoords = new() { BorderSides = ToolStripStatusLabelBorderSides.Left, BorderStyle = Border3DStyle.Etched, Padding = new Padding(5, 0, 5, 0) };


        private ToolStripProgressBar prgExport = new() { Style = ProgressBarStyle.Marquee, Visible = false, MarqueeAnimationSpeed = 30 };
        private ToolStripStatusLabel lblDestLink = new()
        {
            BorderSides = ToolStripStatusLabelBorderSides.Left,
            BorderStyle = Border3DStyle.Etched,
            Padding = new Padding(5, 0, 5, 0),
            IsLink = true,
            Visible = false,
            Text = "Open output",
            // === THÊM DÒNG NÀY ===
            LinkColor = System.Drawing.ColorTranslator.FromHtml("#2563EB")
        };
        #endregion

        #region Constructor and Form Lifecycle

        public MainForm()
        {
            InitializeComponent();
            SetupUi();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            UpdateStatus("Ready. Please open a DOCX or PDF file.");
            UpdateFieldCount();
            UpdateFileName("No file open");
            UpdateCoordinates(0, 0);
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            tplWatcher?.Dispose();
            base.OnFormClosed(e);
        }

        #endregion

        #region UI Setup

        private void SetupUi()
        {
            Text = "PdfSignerStudio (Word Interop + WebView2 + iText7)";
            ClientSize = new Size(1280, 820);
            StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = System.Drawing.ColorTranslator.FromHtml("#F5F5F5");

            topToolstrip = new ToolStrip
            {
                Dock = DockStyle.None,
                GripStyle = ToolStripGripStyle.Hidden,
                // === KÍCH THƯỚC "VỪA PHẢI" ===
                ImageScalingSize = new Size(40, 40),
                LayoutStyle = ToolStripLayoutStyle.HorizontalStackWithOverflow,
                AutoSize = false,
                Stretch = false,
                Height = 60,
                Padding = new Padding(8, 0, 8, 0),
                BackColor = Color.White,
                Renderer = new ToolStripProfessionalRenderer(new CustomColorTable())
            };

            this.Font = new Font("Segoe UI", 9F);

            // ... (Phần còn lại của code giữ nguyên như lần trước)

            btnOpen = new ToolStripButton
            {
                ToolTipText = "Open (Ctrl+O)",
                Image = Properties.Resources.file,
                DisplayStyle = ToolStripItemDisplayStyle.Image,
                ImageScaling = ToolStripItemImageScaling.SizeToFit
            };
            btnOpen.Click += OnOpenFile;

            btnSaveJson = new ToolStripButton
            {
                ToolTipText = "Save Project (Ctrl+S)",
                Image = Properties.Resources.export_json,
                DisplayStyle = ToolStripItemDisplayStyle.Image,
                ImageScaling = ToolStripItemImageScaling.SizeToFit
            };
            btnSaveJson.Click += (_, __) => { _ = SaveJson(); };

            btnLoadJson = new ToolStripButton
            {
                ToolTipText = "Load Project",
                Image = Properties.Resources.import_json,
                DisplayStyle = ToolStripItemDisplayStyle.Image,
                ImageScaling = ToolStripItemImageScaling.SizeToFit
            };
            btnLoadJson.Click += (_, __) => LoadJson();

            btnExport = new ToolStripButton
            {
                ToolTipText = "Export to PDF (Ctrl+E)",
                Image = Properties.Resources.export_pdf,
                DisplayStyle = ToolStripItemDisplayStyle.Image,
                ImageScaling = ToolStripItemImageScaling.SizeToFit
            };
            btnExport.Click += async (_, __) => await ExportPdfAsync();

            btnUndo = new ToolStripButton
            {
                ToolTipText = "Undo (Ctrl+Z)",
                Image = Properties.Resources.undo,
                DisplayStyle = ToolStripItemDisplayStyle.Image,
                ImageScaling = ToolStripItemImageScaling.SizeToFit
            };
            btnUndo.Click += (_, __) => {
                if (_undo.Count > 0)
                {
                    _redo.Push(CloneState(state)); var prev = _undo.Pop(); ApplyState(prev);
                    _isDirty = true;
                }
            };

            btnRedo = new ToolStripButton
            {
                ToolTipText = "Redo (Ctrl+Y)",
                Image = Properties.Resources.redo,
                DisplayStyle = ToolStripItemDisplayStyle.Image,
                ImageScaling = ToolStripItemImageScaling.SizeToFit
            };
            btnRedo.Click += (_, __) => {
                if (_redo.Count > 0)
                {
                    _undo.Push(CloneState(state)); var next = _redo.Pop(); ApplyState(next);
                    _isDirty = true;
                }
            };

            btnGrid = new ToolStripButton
            {
                ToolTipText = "Toggle Grid (G)",
                Image = Properties.Resources.grid,
                DisplayStyle = ToolStripItemDisplayStyle.Image,
                ImageScaling = ToolStripItemImageScaling.SizeToFit
            };
            btnGrid.Click += async (_, __) => { if (web.CoreWebView2 != null) await web.CoreWebView2.ExecuteScriptAsync("if(window.toggleGrid)toggleGrid();"); };

            btnZoomOut = new ToolStripButton
            {
                ToolTipText = "Zoom Out (Ctrl−)",
                Image = Properties.Resources.zoom_out,
                DisplayStyle = ToolStripItemDisplayStyle.Image,
                ImageScaling = ToolStripItemImageScaling.SizeToFit
            };
            btnZoomOut.Click += async (_, __) => { if (web.CoreWebView2 != null) await web.CoreWebView2.ExecuteScriptAsync("zoomOut()"); };

            btnZoomIn = new ToolStripButton
            {
                ToolTipText = "Zoom In (Ctrl+)",
                Image = Properties.Resources.zoom_in,
                DisplayStyle = ToolStripItemDisplayStyle.Image,
                ImageScaling = ToolStripItemImageScaling.SizeToFit
            };
            btnZoomIn.Click += async (_, __) => { if (web.CoreWebView2 != null) await web.CoreWebView2.ExecuteScriptAsync("zoomIn()"); };

            btnTplFolder = new ToolStripButton
            {
                ToolTipText = "Open the templates folder",
                Image = Properties.Resources.opened_folder,
                DisplayStyle = ToolStripItemDisplayStyle.Image,
                ImageScaling = ToolStripItemImageScaling.SizeToFit
            };
            btnTplFolder.Click += (_, __) => { Directory.CreateDirectory(templatesDir); System.Diagnostics.Process.Start("explorer.exe", templatesDir); };


            topToolstrip.Items.AddRange(new ToolStripItem[] {
        btnOpen,
        btnExport,
        new ToolStripSeparator(),
        btnUndo,
        btnRedo,
        new ToolStripSeparator(),
        btnGrid,
        btnZoomIn,
        btnZoomOut,
        new ToolStripSeparator(),
        btnSaveJson,
        btnLoadJson,
        new ToolStripSeparator(),
        btnTplFolder
    });

            foreach (ToolStripItem it in topToolstrip.Items)
            {
                if (it is ToolStripButton)
                {
                    it.Margin = new Padding(3, 3, 3, 3);
                }
                else if (it is ToolStripSeparator)
                {
                    it.Margin = new Padding(6, 0, 6, 0);
                }
            }

            toolHost = new Panel { Dock = DockStyle.Top, Height = topToolstrip.Height, BackColor = Color.White };
            Controls.Add(toolHost);
            toolHost.Controls.Add(topToolstrip);
            this.Load += (_, __) => CenterToolstrip();
            toolHost.Resize += (_, __) => CenterToolstrip();
            topToolstrip.SizeChanged += (_, __) => CenterToolstrip();
            Controls.Add(web);
            Controls.Add(statusBar);

            statusBar.Items.AddRange(new ToolStripItem[] { lblStatus, prgExport, lblDestLink, lblFileName, lblFieldCount, lblCoords });

            lblDestLink.Click += (_, __) =>
            {
                try
                {
                    if (!string.IsNullOrWhiteSpace(lblDestLink.Tag as string))
                    {
                        var path = lblDestLink.Tag!.ToString()!;
                        if (File.Exists(path))
                            System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{path}\"");
                        else if (Directory.Exists(path))
                            System.Diagnostics.Process.Start("explorer.exe", path);
                    }
                }
                catch { /* ignore */ }
            };

            lblStatus.TextAlign = ContentAlignment.MiddleLeft;
            lblFileName.TextAlign = ContentAlignment.MiddleCenter;
            lblFieldCount.TextAlign = ContentAlignment.MiddleCenter;
            lblCoords.TextAlign = ContentAlignment.MiddleRight;

            Load += MainForm_Load;
            RefreshCommandStates();
        }

        // >> THÊM CLASS NHỎ NÀY VÀO BÊN DƯỚI
        public class CustomColorTable : ProfessionalColorTable
        {
            public override Color GripDark => Color.White;
            public override Color GripLight => Color.White;
        }
        #endregion

        #region Status Bar Helpers

        private void UpdateStatus(string message)
        {
            if (IsHandleCreated) BeginInvoke(new Action(() => lblStatus.Text = message));
        }

        private void UpdateFieldCount()
        {
            if (IsHandleCreated) BeginInvoke(new Action(() => lblFieldCount.Text = $"{state.Fields.Count} fields"));
        }

        private void UpdateFileName(string filePath)
        {
            if (IsHandleCreated) BeginInvoke(new Action(() => lblFileName.Text = Path.GetFileName(filePath)));
        }

        private void UpdateCoordinates(float x, float y)
        {
            if (IsHandleCreated) BeginInvoke(new Action(() => lblCoords.Text = $"X: {x:F1}, Y: {y:F1} pt"));
        }

        #endregion

        #region Core Logic: File Open and Processing

        private async void OnOpenFile(object? sender, EventArgs e)
        {
            if (!ConfirmSaveIfDirty()) return;

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
            _ = PushAllFieldsToJs(); // Clear fields in UI

            string outDir = Path.Combine(Path.GetTempPath(), "PdfSignerStudio");
            Directory.CreateDirectory(outDir);

            try
            {
                if (ext == ".docx")
                {
                    SplashForm? splash = null;
                    DimmerForm? dimmer = null;
                    try
                    {
                        dimmer = new DimmerForm { Bounds = this.Bounds };
                        dimmer.Show(this);
                        splash = new SplashForm();
                        splash.Show(this);
                        Application.DoEvents();

                        UpdateStatus("Converting DOCX → PDF with Microsoft Word...");
                        state.SourceDocx = ofd.FileName;
                        state.TempPdf = await RunSTA(() => PdfService.ConvertDocxToPdfWithWord(ofd.FileName, outDir));
                    }
                    finally
                    {
                        splash?.Close();
                        dimmer?.Close();
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
                        // Fallback to original file if copy fails (e.g., permission issue)
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
            cwv2.SetVirtualHostNameToFolderMapping(host, pdfFolder, CoreWebView2HostResourceAccessKind.Allow);

            // *** FIX: Read the static HTML content without replacing URL ***
            var htmlContent = File.ReadAllText(HtmlFilePath());

            web.CoreWebView2.NavigationCompleted -= OnWebReady;
            web.CoreWebView2.NavigationCompleted += OnWebReady;

            // *** FIX: Navigate to the static HTML content. The PDF will be loaded via JS call later. ***
            web.CoreWebView2.NavigateToString(htmlContent);

            UpdateFileName(ofd.FileName);
            UpdateStatus("Ready. Drag, drop, nudge, snap, rename inline, flip pages with mouse/PageUp-Down.");

            _isDirty = false;
        }

        private async void OnWebReady(object? sender, CoreWebView2NavigationCompletedEventArgs e)
        {
            web.CoreWebView2.NavigationCompleted -= OnWebReady;

            // *** FIX: This event fires after HTML is loaded. Now it's safe to tell JS to load the PDF. ***

            // 1. Construct the virtual PDF URI
            var host = "files.local";
            var pdfUri = $"https://{host}/{Path.GetFileName(state.TempPdf!)}";

            // 2. Call the new JavaScript function and pass the URI
            await web.CoreWebView2.ExecuteScriptAsync($"initializePdfViewer('{pdfUri}');");

            // 3. Load templates and fields as before
            LoadTemplates();
            await PushTemplatesToJs();
            await PushAllFieldsToJs();
            SetupTplWatcher();
            UpdateFieldCount();
        }

        #endregion

        #region WebView2 Communication (JS -> C#)

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
                            int page = root.GetProperty("page").GetInt32();
                            _currentPage = page;
                            PushFieldsToJs(page);
                            break;
                        }
                    case "addField":
                        PushUndo();
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

                            // === LOGIC MỚI: Tự động xử lý trùng tên ===
                            string baseName = name.Trim();
                            string finalName = baseName;
                            int idx = 1;
                            while (state.Fields.Any(f => f.Name.Equals(finalName, StringComparison.OrdinalIgnoreCase)))
                            {
                                finalName = $"{baseName}_{idx++}";
                            }
                            // =======================================

                            state.Fields.Add(new FormFieldDef(finalName, "signature", page, new RectFpt(x, y, w, h), req));
                            UpdateStatus($"Added {finalName} on page {page}");
                            _currentPage = page;
                            PushFieldsToJs(page);
                            UpdateFieldCount();
                            _ = PushAllFieldsToJs(); // Update the list of all fields
                            break;
                        }
                    case "updateField":
                        PushUndo();
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
                                _currentPage = page;
                                PushFieldsToJs(page);
                                _ = PushAllFieldsToJs(); // Update list
                            }
                            break;
                        }
                    case "deleteField":
                        PushUndo();
                        {
                            string id = root.GetProperty("id").GetString()!;
                            int page = root.GetProperty("page").GetInt32();
                            state.Fields.RemoveAll(f => f.Id == id);
                            UpdateStatus("Field deleted.");
                            _currentPage = page;
                            PushFieldsToJs(page);
                            UpdateFieldCount();
                            _ = PushAllFieldsToJs(); // Update list
                            break;
                        }
                    case "renameField":
                        PushUndo();
                        {
                            string id = root.GetProperty("id").GetString()!;
                            string newName = root.GetProperty("name").GetString() ?? "";
                            int page = root.GetProperty("page").GetInt32();
                            if (string.IsNullOrWhiteSpace(newName)) break;

                            string baseName = newName.Trim();
                            string name = baseName;
                            int idx = 1;
                            while (state.Fields.Any(f => f.Name.Equals(name, StringComparison.OrdinalIgnoreCase) && f.Id != id))
                                name = $"{baseName}_{idx++}";

                            var f = state.Fields.FirstOrDefault(t => t.Id == id);
                            if (f != null)
                            {
                                state.Fields[state.Fields.IndexOf(f)] = f with { Name = name };
                                _currentPage = page;
                                PushFieldsToJs(page);
                                _ = PushAllFieldsToJs(); // Update list
                            }
                            break;
                        }
                    case "toggleRequired":
                        PushUndo();
                        {
                            string id = root.GetProperty("id").GetString()!;
                            int page = root.GetProperty("page").GetInt32();
                            var f = state.Fields.FirstOrDefault(t => t.Id == id);
                            if (f != null)
                            {
                                state.Fields[state.Fields.IndexOf(f)] = f with { Required = !f.Required };
                                _currentPage = page;
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
                                foreach (var c in Path.GetInvalidFileNameChars()) s = s.Replace(c, '_');
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
                                foreach (var c in Path.GetInvalidFileNameChars()) s = s.Replace(c, '_');
                                return s.Trim();
                            }

                            var path = Path.Combine(templatesDir, Safe(name) + ".json");
                            if (File.Exists(path)) File.Delete(path);

                            LoadTemplates();
                            _ = PushTemplatesToJs();
                            UpdateStatus($"Deleted template: {name}");
                            break;
                        }
                    case "undo":
                        {
                            if (_undo.Count > 0)
                            {
                                _redo.Push(CloneState(state));
                                var prev = _undo.Pop();
                                ApplyState(prev);
                                _isDirty = true;
                            }
                            break;
                        }
                    case "redo":
                        {
                            if (_redo.Count > 0)
                            {
                                _undo.Push(CloneState(state));
                                var next = _redo.Pop();
                                ApplyState(next);
                                _isDirty = true;
                            }
                            break;
                        }

                }
            }
            catch
            {
                // Ignore parsing errors
            }
        }

        #endregion

        #region WebView2 Communication (C# -> JS)

        private async void PushFieldsToJs(int page)
        {
            if (web.CoreWebView2 == null) return;
            // Ensure every field has a non-null ID
            for (int i = 0; i < state.Fields.Count; i++)
            {
                if (string.IsNullOrEmpty(state.Fields[i].Id))
                    state.Fields[i] = state.Fields[i] with { };
            }

            var list = state.Fields.Where(f => f.Page == page)
                .Select(f => new { id = f.Id, name = f.Name, x = f.Rect.X, y = f.Rect.Y, w = f.Rect.W, h = f.Rect.H, required = f.Required });

            string json = JsonSerializer.Serialize(list);
            await web.CoreWebView2.ExecuteScriptAsync($"setFields({json});");
        }

        private async Task PushAllFieldsToJs()
        {
            if (web.CoreWebView2 == null) return;

            var allFields = state.Fields
                .OrderBy(f => f.Page)
                .ThenBy(f => f.Name)
                .Select(f => new { id = f.Id, name = f.Name, page = f.Page });

            string json = JsonSerializer.Serialize(allFields);
            await web.CoreWebView2.ExecuteScriptAsync($"setAddedFields({json});");
        }

        private async Task PushTemplatesToJs()
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

        #endregion

        #region Template Management

        private void SetupTplWatcher()
        {
            try
            {
                tplWatcher?.Dispose();
                Directory.CreateDirectory(templatesDir);
                tplWatcher = new FileSystemWatcher(templatesDir, "*.json")
                {
                    IncludeSubdirectories = false,
                    NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite | NotifyFilters.Size
                };
                tplWatcher.Changed += OnTplChanged;
                tplWatcher.Created += OnTplChanged;
                tplWatcher.Deleted += OnTplChanged;
                tplWatcher.Renamed += OnTplRenamed;
                tplWatcher.EnableRaisingEvents = true;
            }
            catch { /* Ignore watcher setup errors */ }
        }

        private async void OnTplChanged(object? sender, FileSystemEventArgs e)
        {
            try
            {
                await Task.Delay(250); // Debounce
                if (IsDisposed) return;
                BeginInvoke(new Action(async () =>
                {
                    LoadTemplates();
                    await PushTemplatesToJs();
                    UpdateStatus("Templates reloaded.");
                }));
            }
            catch { /* Ignore */ }
        }

        private void OnTplRenamed(object? s, RenamedEventArgs e) => OnTplChanged(s, e);

        private void LoadTemplates()
        {
            templates.Clear();
            Directory.CreateDirectory(templatesDir);

            // Create demo templates if the directory is empty
            if (!Directory.EnumerateFiles(templatesDir, "*.json").Any())
            {
                var demo1 = new TemplateDef("Signature 120×60", new List<TemplateField> { new TemplateField("Signature", 120, 60, true, 0, 0) });
                var demo2 = new TemplateDef("Director + Accountant", new List<TemplateField>
                {
                    new TemplateField("Director", 140, 70, true, 0, 0),
                    new TemplateField("Accountant", 140, 70, true, 160, 0),
                });
                File.WriteAllText(Path.Combine(templatesDir, "Signature_120x60.json"), JsonSerializer.Serialize(demo1, new JsonSerializerOptions { WriteIndented = true }));
                File.WriteAllText(Path.Combine(templatesDir, "Director_Accountant.json"), JsonSerializer.Serialize(demo2, new JsonSerializerOptions { WriteIndented = true }));
            }

            foreach (var f in Directory.EnumerateFiles(templatesDir, "*.json"))
            {
                try
                {
                    var t = JsonSerializer.Deserialize<TemplateDef>(File.ReadAllText(f));
                    if (t?.Items != null && t.Items.Count > 0)
                        templates.Add(t);
                }
                catch { /* Ignore malformed JSON */ }
            }
        }

        #endregion

        #region Project State: Save/Load JSON and Export PDF


        private bool SaveJson()
        {
            using var sfd = new SaveFileDialog { Filter = "JSON (*.json)|*.json" };
            if (sfd.ShowDialog() != DialogResult.OK) return false;

            var options = new JsonSerializerOptions { WriteIndented = true };
            File.WriteAllText(sfd.FileName, JsonSerializer.Serialize(state, options));
            UpdateStatus("Saved JSON project.");
            _isDirty = false;
            return true;
        }



        private bool ConfirmSaveIfDirty()
        {
            if (!_isDirty) return true;
            var r = MessageBox.Show("Save changes?", "Unsaved changes", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1);
            if (r == DialogResult.Yes) return SaveJson();
            if (r == DialogResult.No) return true;
            return false;
        }

        private async void LoadJson()
        {
            if (!ConfirmSaveIfDirty()) return;

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

                // Ensure all fields have an ID after loading
                for (int i = 0; i < state.Fields.Count; i++)
                    if (string.IsNullOrEmpty(state.Fields[i].Id))
                        state.Fields[i] = state.Fields[i] with { };

                UpdateFileName(ofd.FileName);
                UpdateStatus("Loaded JSON project successfully.");

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
                    cwv2.SetVirtualHostNameToFolderMapping(host, pdfFolder, CoreWebView2HostResourceAccessKind.Allow);

                    var htmlContent = File.ReadAllText(HtmlFilePath());

                    web.CoreWebView2.NavigationCompleted -= OnWebReady;
                    web.CoreWebView2.NavigationCompleted += OnWebReady;
                    web.CoreWebView2.NavigateToString(htmlContent);
                }
                else
                {
                    await PushAllFieldsToJs(); // Push fields, but don't load PDF preview
                    MessageBox.Show("JSON file does not contain a valid path to a PDF (TempPdf). Please open a DOCX/PDF first, or edit the 'TempPdf' path in the JSON.", "Missing PDF", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("JSON load failed: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            _isDirty = false;
        }

        // === Export UI helpers ===
        private void BeginExportUi(string msg)
        {
            if (!IsHandleCreated) return;
            BeginInvoke(new Action(() =>
            {
                lblStatus.Text = msg;
                prgExport.Visible = true;
                prgExport.MarqueeAnimationSpeed = 30;
                lblDestLink.Visible = false;
                lblDestLink.Tag = null;
            }));
        }

        private void StepExportUi(string msg)
        {
            if (!IsHandleCreated) return;
            BeginInvoke(new Action(() =>
            {
                lblStatus.Text = msg;
            }));
        }

        private void EndExportUi(string finalMsg, string? outputPath)
        {
            if (!IsHandleCreated) return;
            BeginInvoke(new Action(() =>
            {
                lblStatus.Text = finalMsg;
                prgExport.Visible = false;
                prgExport.MarqueeAnimationSpeed = 0;

                if (!string.IsNullOrWhiteSpace(outputPath))
                {
                    lblDestLink.Text = "Open output";
                    lblDestLink.Tag = outputPath;
                    lblDestLink.Visible = true;
                }
                else
                {
                    lblDestLink.Visible = false;
                }
            }));
        }


        private async Task ExportPdfAsync()
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
                BeginExportUi("Exporting PDF…");
                StepExportUi("Preparing fields…");
                var fieldsCopy = state.Fields.ToList();

                StepExportUi("Writing PDF…");
                await Task.Run(() =>
                {
                    PdfService.AddSignatureFields(state.TempPdf!, sfd.FileName, fieldsCopy);
                });

                EndExportUi("Export complete.", sfd.FileName);
                try { System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{sfd.FileName}\""); } catch { /* ignore */ }
            }
            catch (IOException)
            {
                EndExportUi("Export failed: file is in use.", null);
                MessageBox.Show(
                    $"Cannot save the file.\n\nThe file '{Path.GetFileName(sfd.FileName)}' might be open in another program (like Adobe Reader, Chrome, etc.).\n\nPlease close that file and try again.",
                    "File Write Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                EndExportUi("Export failed.", null);
                MessageBox.Show("An error occurred during export:\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExportPdf()
        {
            _ = ExportPdfAsync();
        }


        #endregion


        // ===== Undo/Redo helpers =====
        private ProjectState CloneState(ProjectState s)
        {
            var copy = new ProjectState
            {
                SourceDocx = s.SourceDocx,
                TempPdf = s.TempPdf
            };
            foreach (var f in s.Fields)
                copy.Fields.Add(f with { });
            return copy;
        }

        private
        void PushUndo()
        {
            _undo.Push(CloneState(state));
            _redo.Clear();
            _isDirty = true;
            RefreshCommandStates();
        }


        private void ApplyState(ProjectState s)
        {
            state = CloneState(s);
            UpdateFieldCount();
            _ = PushAllFieldsToJs();
            PushFieldsToJs(_currentPage);
            RefreshCommandStates();
        }

        #region Helper Methods

        private async Task EnsureWebReady()
        {
            if (web.CoreWebView2 == null)
                await web.EnsureCoreWebView2Async();
            RefreshCommandStates();
        }

        private static string HtmlFilePath()
        {
            string path = Path.Combine(AppContext.BaseDirectory, "Web", "index.html");
            if (!File.Exists(path))
                throw new FileNotFoundException("Web\\index.html not found. Please create the file and set its 'Copy to Output Directory' property to 'Copy if newer'.", path);
            return path;
        }

        // This method is no longer needed as we navigate directly to string
        // private string BuildPdfHtml(string pdfFileUri) { ... }

        private static Task<T> RunSTA<T>(Func<T> func)
        {
            var tcs = new TaskCompletionSource<T>();
            var th = new Thread(() =>
            {
                try
                {
                    tcs.SetResult(func());
                }
                catch (Exception ex)
                {
                    tcs.SetException(ex);
                }
            });
            th.SetApartmentState(ApartmentState.STA);
            th.Start();
            return tcs.Task;
        }

        #endregion

        private void CenterToolstrip()
        {
            if (topToolstrip == null || toolHost == null) return;

            // lấy kích thước thực của các item
            int w = topToolstrip.PreferredSize.Width;
            int h = topToolstrip.PreferredSize.Height;

            // ép width/h đúng bằng content để khỏi wrap
            topToolstrip.Width = w;
            topToolstrip.Height = h;

            int x = Math.Max(0, (toolHost.ClientSize.Width - w) / 2);
            int y = (toolHost.ClientSize.Height - h) / 2;

            topToolstrip.Location = new Point(x, y);
        }




        private void RefreshCommandStates()
        {
            try
            {
                if (btnUndo != null) btnUndo.Enabled = _undo != null && _undo.Count > 0;
                if (btnRedo != null) btnRedo.Enabled = _redo != null && _redo.Count > 0;
                if (btnGrid != null) btnGrid.Enabled = web != null && web.CoreWebView2 != null;
            }
            catch { /* ignore */ }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (!ConfirmSaveIfDirty())
            {
                e.Cancel = true;
                return;
            }
            base.OnFormClosing(e);
        }
    }
}
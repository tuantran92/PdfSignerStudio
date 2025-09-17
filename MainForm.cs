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

        // Template library
        List<TemplateDef> templates = new();
        string templatesDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "PdfSignerStudio", "Templates");

        // Auto-reload watcher
        FileSystemWatcher? tplWatcher;

        public MainForm()
        {
            InitializeComponent();
            SetupUi();
        }

        private void SetupUi()
        {
            Text = "PdfSignerStudio (Word Interop + WebView2 + iText7)";
            ClientSize = new Size(1280, 820);
            StartPosition = FormStartPosition.CenterScreen;

            topBar.Controls.AddRange(new Control[]
            {
                btnOpenDocx, pageBox, btnZoomOut, btnZoomIn,
                btnSaveJson, btnLoadJson, btnExport, btnTplFolder, info
            });
            btnOpenDocx.Left = 8;
            pageBox.Left = 120;
            btnZoomOut.Left = 210; btnZoomOut.Width = 30;
            btnZoomIn.Left = 245; btnZoomIn.Width = 30;
            btnSaveJson.Left = 290;
            btnLoadJson.Left = 390;
            btnExport.Left = 490;
            btnTplFolder.Left = 590;
            info.Left = 700; info.Top = 12;

            Controls.Add(web);
            Controls.Add(topBar);

            btnOpenDocx.Click += OnOpenFile; // <— đổi handler
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
        }

        // Run Interop Word in STA thread
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

        // === NEW: mở DOCX hoặc PDF ===
        async void OnOpenFile(object? sender, EventArgs e)
        {
            using var ofd = new OpenFileDialog
            {
                Filter = "Word (*.docx)|*.docx|PDF (*.pdf)|*.pdf",
                Title = "Open Word/PDF"
            };
            if (ofd.ShowDialog() != DialogResult.OK) return;

            string ext = Path.GetExtension(ofd.FileName).ToLowerInvariant();
            state = new ProjectState();
            string outDir = Path.Combine(Path.GetTempPath(), "PdfSignerStudio");
            Directory.CreateDirectory(outDir);

            try
            {
                if (ext == ".docx")
                {
                    info.Text = "Converting DOCX → PDF with Microsoft Word...";
                    state.SourceDocx = ofd.FileName;
                    state.TempPdf = await RunSTA(() =>
                        PdfService.ConvertDocxToPdfWithWord(ofd.FileName, outDir));
                }
                else // .pdf
                {
                    info.Text = "Loading PDF…";
                    state.SourceDocx = null;
                    // Copy qua temp để tránh bị khóa/di chuyển nguồn
                    string dest = Path.Combine(outDir, Path.GetFileName(ofd.FileName));
                    try
                    {
                        File.Copy(ofd.FileName, dest, overwrite: true);
                        state.TempPdf = dest;
                    }
                    catch
                    {
                        // Nếu copy fail (VD khác ổ) thì cứ dùng file gốc
                        state.TempPdf = ofd.FileName;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Open failed: " + ex.Message);
                info.Text = "Open failed.";
                return;
            }

            // Nạp preview PDF
            info.Text = "Loading preview...";
            await EnsureWebReady();

            web.CoreWebView2.WebMessageReceived -= WebMessageReceived;
            web.CoreWebView2.WebMessageReceived += WebMessageReceived;

            var pdfFolder = Path.GetDirectoryName(state.TempPdf!)!;
            web.CoreWebView2.SetVirtualHostNameToFolderMapping(
                "app", pdfFolder, CoreWebView2HostResourceAccessKind.Allow);

            var pdfUri = $"https://app/{Path.GetFileName(state.TempPdf!)}";
            var html = BuildPdfHtml(pdfUri);

            web.CoreWebView2.NavigationCompleted -= OnWebReady;
            web.CoreWebView2.NavigationCompleted += OnWebReady;

            web.CoreWebView2.NavigateToString(html);

            info.Text = "Ready. Kéo–thả, nudge, snap, rename inline, lật trang bằng chuột/PageUp-Down.";
        }

        // Gọi sau mỗi lần NavigateToString hoàn tất
        private async void OnWebReady(object? sender, CoreWebView2NavigationCompletedEventArgs e)
        {
            web.CoreWebView2.NavigationCompleted -= OnWebReady;

            LoadTemplates();
            await PushTemplatesToJs();
            SetupTplWatcher();

            if (pageBox.SelectedIndex >= 0)
            {
                int page = pageBox.SelectedIndex + 1;
                PushFieldsToJs(page);
            }
        }

        // Auto-reload watcher
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
            catch { /* ignore */ }
        }

        async void OnTplChanged(object? sender, FileSystemEventArgs e)
        {
            try
            {
                await Task.Delay(250); // debounce
                if (IsDisposed) return;
                BeginInvoke(new Action(async () =>
                {
                    LoadTemplates();
                    await PushTemplatesToJs();
                    info.Text = "Templates reloaded.";
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

        // nhận message từ HTML
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
                            info.Text = $"Added {name} on page {page}";
                            PushFieldsToJs(page);
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
                            }
                            break;
                        }
                    case "deleteField":
                        {
                            string id = root.GetProperty("id").GetString()!;
                            int page = root.GetProperty("page").GetInt32();
                            state.Fields.RemoveAll(f => f.Id == id);
                            PushFieldsToJs(page);
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

            // ensure Id exists (các JSON cũ)
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

        // ====== Templates ======
        void LoadTemplates()
        {
            templates.Clear();
            Directory.CreateDirectory(templatesDir);

            // Seed demo nếu trống
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
                catch { /* skip invalid */ }
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
            info.Text = "Saved JSON.";
        }

        async void LoadJson()
        {
            using var ofd = new OpenFileDialog { Filter = "JSON (*.json)|*.json" };
            if (ofd.ShowDialog() != DialogResult.OK) return;

            try
            {
                state = JsonSerializer.Deserialize<ProjectState>(File.ReadAllText(ofd.FileName)) ?? new ProjectState();

                // đảm bảo có Id
                for (int i = 0; i < state.Fields.Count; i++)
                    if (string.IsNullOrEmpty(state.Fields[i].Id))
                        state.Fields[i] = state.Fields[i] with { };

                info.Text = "Loaded JSON.";

                if (!string.IsNullOrEmpty(state.TempPdf) && File.Exists(state.TempPdf))
                {
                    await EnsureWebReady();

                    web.CoreWebView2.WebMessageReceived -= WebMessageReceived;
                    web.CoreWebView2.WebMessageReceived += WebMessageReceived;

                    var pdfFolder = Path.GetDirectoryName(state.TempPdf!)!;
                    web.CoreWebView2.SetVirtualHostNameToFolderMapping(
                        "app", pdfFolder, CoreWebView2HostResourceAccessKind.Allow);

                    var pdfUri = $"https://app/{Path.GetFileName(state.TempPdf!)}";
                    var html = BuildPdfHtml(pdfUri);

                    web.CoreWebView2.NavigationCompleted -= OnWebReady;
                    web.CoreWebView2.NavigationCompleted += OnWebReady;

                    web.CoreWebView2.NavigateToString(html);
                }
                else
                {
                    MessageBox.Show("JSON chưa có đường dẫn PDF hợp lệ. Hãy Open một DOCX/PDF trước, hoặc chỉnh lại 'TempPdf' trong JSON.");
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
            if (state.TempPdf == null) { MessageBox.Show("Chưa có PDF."); return; }
            using var sfd = new SaveFileDialog { Filter = "PDF (*.pdf)|*.pdf" };
            if (sfd.ShowDialog() != DialogResult.OK) return;

            try
            {
                PdfService.AddSignatureFields(state.TempPdf, sfd.FileName, state.Fields);
                info.Text = $"Exported: {sfd.FileName}";
                MessageBox.Show("Xuất PDF thành công!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Export failed:\n" + ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async Task EnsureWebReady()
        {
            if (web.CoreWebView2 == null)
                await web.EnsureCoreWebView2Async();
        }
        private static string HtmlFilePath()
        {
            // trỏ tới Web\index.html nằm cạnh exe
            return Path.Combine(AppContext.BaseDirectory, "Web", "index.html");
        }

        private string BuildPdfHtml(string pdfFileUri)
        {
            var path = HtmlFilePath();
            if (!File.Exists(path))
                throw new FileNotFoundException("Không tìm thấy Web\\index.html. Hãy tạo file và đặt Copy to Output = Copy if newer.", path);

            var tpl = File.ReadAllText(path);
            return tpl.Replace("__PDF_URL__", pdfFileUri);
        }

    }
}

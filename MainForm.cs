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

        // ===== Template library =====
        List<TemplateDef> templates = new();
        string templatesDir = Path.Combine(AppContext.BaseDirectory, "Template");
        FileSystemWatcher? tplWatcher;

        // --- Panel bên phải đã được XÓA BỎ ---

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

            // --- Code thiết lập panel bên phải đã được XÓA BỎ ---

            // Đặt lại thứ tự controls như ban đầu
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

            Load += (_, __) => { CenterToolbar(); PositionInfo(); };
            Resize += (_, __) => { CenterToolbar(); PositionInfo(); };
            topBar.Resize += (_, __) => { CenterToolbar(); PositionInfo(); };
            info.SizeChanged += (_, __) => PositionInfo();
        }

        // ===== MODIFIED: Phương thức mới để đẩy TOÀN BỘ danh sách fields xuống JS =====
        async Task PushAllFieldsToJs()
        {
            if (web.CoreWebView2 == null) return;

            // Lấy tất cả các field, sắp xếp và CHỌN THÊM ID
            var allFields = state.Fields
                .OrderBy(f => f.Page)
                .ThenBy(f => f.Name)
                .Select(f => new { id = f.Id, name = f.Name, page = f.Page }); // <--- THÊM "id = f.Id" VÀO ĐÂY

            string json = JsonSerializer.Serialize(allFields);
            // Gọi hàm JavaScript `setAddedFields` (sẽ được tạo ở Bước 2)
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
            _ = PushAllFieldsToJs(); // MODIFIED: Cập nhật danh sách (rỗng) trên UI web
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
                else
                {
                    info.Text = "Loading PDF…";
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
                info.Text = "Open failed.";
                return;
            }

            info.Text = "Loading preview...";
            await EnsureWebReady();

            web.CoreWebView2.WebMessageReceived -= WebMessageReceived;
            web.CoreWebView2.WebMessageReceived += WebMessageReceived;

            var host = "files.local";
            var pdfFolder = Path.GetDirectoryName(state.TempPdf!)!;
            pdfFolder = Path.GetFullPath(pdfFolder);

            if (!Directory.Exists(pdfFolder))
                throw new DirectoryNotFoundException(pdfFolder);

            var cwv2 = web.CoreWebView2;
            try { cwv2.ClearVirtualHostNameToFolderMapping(host); } catch { /* ignore */ }
            cwv2.SetVirtualHostNameToFolderMapping(
                host, pdfFolder, CoreWebView2HostResourceAccessKind.Allow);

            var pdfUri = $"https://{host}/{Path.GetFileName(state.TempPdf!)}";
            var html = BuildPdfHtml(pdfUri);

            web.CoreWebView2.NavigationCompleted -= OnWebReady;
            web.CoreWebView2.NavigationCompleted += OnWebReady;

            web.CoreWebView2.NavigateToString(html);

            info.Text = "Ready. Kéo–thả, nudge, snap, rename inline, lật trang bằng chuột/PageUp-Down.";
        }

        private async void OnWebReady(object? sender, CoreWebView2NavigationCompletedEventArgs e)
        {
            web.CoreWebView2.NavigationCompleted -= OnWebReady;

            LoadTemplates();
            await PushTemplatesToJs();
            await PushAllFieldsToJs(); // MODIFIED: Đẩy danh sách fields xuống khi web sẵn sàng
            SetupTplWatcher();

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
                            _ = PushAllFieldsToJs(); // MODIFIED: Cập nhật lại toàn bộ danh sách
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
                                _ = PushAllFieldsToJs(); // MODIFIED: Cập nhật lại toàn bộ danh sách
                            }
                            break;
                        }
                    case "deleteField":
                        {
                            string id = root.GetProperty("id").GetString()!;
                            int page = root.GetProperty("page").GetInt32();
                            state.Fields.RemoveAll(f => f.Id == id);
                            PushFieldsToJs(page);
                            _ = PushAllFieldsToJs(); // MODIFIED: Cập nhật lại toàn bộ danh sách
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
                                _ = PushAllFieldsToJs(); // MODIFIED: Cập nhật lại toàn bộ danh sách
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

                    // ========= Template CRUD ========= (Không thay đổi)
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
                            info.Text = $"Saved template: {name}";
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
                            info.Text = $"Deleted template: {name}";
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

                // MODIFIED: Cập nhật danh sách trên UI web sau khi load JSON
                // (sẽ được gọi trong OnWebReady sau khi điều hướng lại)

                info.Text = "Loaded JSON.";

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
                    try { cwv2.ClearVirtualHostNameToFolderMapping(host); } catch { /* ignore */ }
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
                    await PushAllFieldsToJs(); // Cập nhật danh sách fields dù không có PDF
                    MessageBox.Show(
                        "JSON chưa có đường dẫn PDF hợp lệ (TempPdf). Hãy Open một DOCX/PDF trước, hoặc chỉnh lại 'TempPdf' trong JSON.",
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
                MessageBox.Show("Chưa có file PDF nào được mở.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using var sfd = new SaveFileDialog { Filter = "PDF (*.pdf)|*.pdf", Title = "Lưu file PDF" };
            if (sfd.ShowDialog() != DialogResult.OK) return;

            try
            {
                PdfService.AddSignatureFields(state.TempPdf, sfd.FileName, state.Fields);
                info.Text = $"Exported: {sfd.FileName}";
                MessageBox.Show($"Xuất file PDF thành công!\nĐã lưu tại: {sfd.FileName}", "Hoàn tất", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (IOException ioEx)
            {
                MessageBox.Show(
                    $"Không thể lưu file.\n\nFile '{Path.GetFileName(sfd.FileName)}' có thể đang được mở trong một chương trình khác (như Adobe Reader, Chrome...).\n\nVui lòng đóng file đó và thử lại.",
                    "Lỗi Ghi File",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
                info.Text = "Export failed: File in use.";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Đã có lỗi xảy ra trong quá trình xuất file:\n" + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                info.Text = "Export failed.";
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
                throw new FileNotFoundException("Không tìm thấy Web\\index.html. Hãy tạo file và đặt Copy to Output = Copy if newer.", path);

            var tpl = File.ReadAllText(path);
            return tpl.Replace("__PDF_URL__", pdfFileUri);
        }
    }
}
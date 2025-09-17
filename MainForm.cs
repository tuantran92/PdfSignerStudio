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

        // ========= RAW HTML template (Replace __PDF_URL__) =========
        private static readonly string HtmlTemplate = """
<!DOCTYPE html><html>
<head>
<meta charset="utf-8"/>
<meta http-equiv="X-UA-Compatible" content="IE=edge"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<style>
  html, body { margin:0; padding:0; height:100%; overflow:hidden; background:#fafafa; user-select:none; }
  #container { position:relative; width:100%; height:100%; overflow:auto; }
  #pagewrap { position:relative; margin:24px auto; width:fit-content; }
  #pdfCanvas { display:block; background:#fff; box-shadow:0 0 20px rgba(0,0,0,.08); }
  #overlay { position:absolute; left:0; top:0; pointer-events:auto; }

  .box { position:absolute; border:2px solid red; box-sizing:border-box; }
  .box.req { border-color:red; }
  .box.nreq { border-color:#888; border-style:dashed; }
  .box.sel { outline:2px solid #1976d2; }
  .handle { position:absolute; width:10px; height:10px; background:#1976d2; right:-6px; bottom:-6px; cursor:nwse-resize; }

  .gridbg {
    background-image:
      linear-gradient(to right, rgba(0,0,0,.06) 1px, transparent 1px),
      linear-gradient(to bottom, rgba(0,0,0,.06) 1px, transparent 1px);
    background-size: var(--grid) var(--grid), var(--grid) var(--grid);
  }

  /* Templates panel (left) */
  #tplbar {
    position:fixed; left:12px; top:60px; width:200px; bottom:12px; background:#fff;
    box-shadow:0 2px 12px rgba(0,0,0,.15); border-radius:10px; padding:10px; overflow:auto; z-index:9999;
    font-family:system-ui,Segoe UI,Roboto,Arial; font-size:12px;
  }
  #tplbar h3 { margin:0 0 8px 0; font-size:13px; }
  .tpl { border:1px solid #ddd; border-radius:8px; padding:8px; margin:6px 0; cursor:grab; background:#fdfdfd; }
  .tpl:hover { background:#f7faff; }

  /* Thumbnails panel (right) */
  #thumbbar {
    position:fixed; right:12px; top:60px; width:180px; bottom:12px; background:#fff;
    box-shadow:0 2px 12px rgba(0,0,0,.15); border-radius:10px; padding:10px; overflow:auto; z-index:9999;
    font-family:system-ui,Segoe UI,Roboto,Arial; font-size:12px;
  }
  #thumbbar h3 { margin:0 0 8px 0; font-size:13px; }
  .thumb {
    border:1px solid #ddd; border-radius:8px; padding:6px; margin:8px 0; cursor:pointer; background:#fff;
    display:flex; flex-direction:column; align-items:center; gap:6px;
  }
  .thumb canvas, .thumb img {
    width:140px; height:auto; border:1px solid #eee; border-radius:4px;
  }
  .thumb .pnum { color:#555; font-size:12px; }
  .thumb.cur { outline:2px solid #1976d2; }
  .thumb:hover { background:#f7faff; }
</style>
</head>
<body>
  <div id="tplbar"><h3>Templates</h3><div id="tpllist">Loading…</div></div>
  <div id="thumbbar"><h3>Pages</h3><div id="thumblist">Loading…</div></div>

  <div id="container">
    <div id="pagewrap">
      <canvas id="pdfCanvas"></canvas>
      <div id="overlay" tabindex="0" class="gridbg"></div>
    </div>
  </div>

<script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>
<script>
  const pdfjsLib = window['pdfjs-dist/build/pdf'];
  pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

  const pdfUrl = '__PDF_URL__';
  let pdfDoc = null, pageNum = 1, scale = 1.25;
  let viewport = null;
  const canvas = document.getElementById('pdfCanvas');
  const ctx = canvas.getContext('2d');
  const overlay = document.getElementById('overlay');
  const container = document.getElementById('container');
  const thumbList = document.getElementById('thumblist');
  let thumbCanvases = [];

  let lastFlipTs = 0;                  // debounce lật trang
  let flipScrollTarget = null;         // 'top' | 'bottom' | null

  // ====== Grid / snap ======
  let gridPt = 8;                      // 1 ô lưới = 8pt
  let snapTolPx = 5;
  const MIN_GRID_CELLS = 3;            // tối thiểu 3x3 ô lưới để tạo field

  function ptToPx(v){ return v * scale; }
  function pxToPt(v){ return v / scale; }
  function pageHeightPt(){ return viewport.height / scale; }
  function gridPx(){ return Math.max(4, Math.round(ptToPx(gridPt))); }
  function applyGridBg(){ overlay.style.setProperty('--grid', gridPx() + 'px'); }

  function snapValPx(v, candidates){
    let best = v;
    for (const c of candidates){ if (Math.abs(v - c) <= snapTolPx){ best = c; break; } }
    return best;
  }
  function guideCandidatesPx(includeBoxes=true){
    const xCands = [], yCands = [];
    const w = canvas.width, h = canvas.height;
    xCands.push(0, w/2, w);
    yCands.push(0, h/2, h);
    if (includeBoxes){
      document.querySelectorAll('.box').forEach(b => {
        const x = b.offsetLeft, y = b.offsetTop, bw = b.offsetWidth, bh = b.offsetHeight;
        xCands.push(x, x + bw/2, x + bw);
        yCands.push(y, y + bh/2, y + bh);
      });
    }
    return { xCands, yCands };
  }
  function snapRectPx(left, top, width, height){
    const {xCands, yCands} = guideCandidatesPx(true);
    const sx1 = snapValPx(left, xCands),   sx2 = snapValPx(left + width, xCands);
    const sy1 = snapValPx(top, yCands),    sy2 = snapValPx(top + height, yCands);
    let l = Math.abs(sx1 - left) <= Math.abs(sx2 - (left+width)) ? sx1 : (sx2 - width);
    let t = Math.abs(sy1 - top)  <= Math.abs(sy2 - (top+height)) ? sy1 : (sy2 - height);
    return { left:l, top:t, width, height };
  }

  // ====== State & selection ======
  let fields = [];                      // [{id,name,x,y,w,h,required}]
  let selectedId = null;
  let dragging = null;                  // {mode:'move'|'resize'|'draw', sx,sy, startRect, moved?}
  let editingId = null;                 // đang edit tên field nào

  function post(msg){ chrome.webview.postMessage(JSON.stringify(msg)); }

  // ====== Render page ======
  async function renderPage(num){
    const page = await pdfDoc.getPage(num);
    viewport = page.getViewport({ scale });
    canvas.width = viewport.width;  canvas.height = viewport.height;
    overlay.style.width = canvas.width + 'px'; overlay.style.height = canvas.height + 'px';
    overlay.style.left = '0px'; overlay.style.top = '0px';
    applyGridBg();

    await page.render({ canvasContext: ctx, viewport }).promise;

    // Đặt lại scroll theo hướng lật
    if (flipScrollTarget === 'top') {
      container.scrollTop = 1;
    } else if (flipScrollTarget === 'bottom') {
      container.scrollTop = Math.max(0, container.scrollHeight - container.clientHeight - 1);
    }
    flipScrollTarget = null;

    post({ type:'meta', numPages: pdfDoc.numPages, page:num });
    markCurrentThumb(pageNum);
    redraw();
  }

  // ====== Thumbnails ======
  async function buildThumbs() {
    thumbList.innerHTML = '';
    thumbCanvases = [];
    const desiredWidth = 140;

    for (let i = 1; i <= pdfDoc.numPages; i++) {
      const page = await pdfDoc.getPage(i);
      const vp0 = page.getViewport({ scale: 1 });
      const scaleT = desiredWidth / vp0.width;
      const vp = page.getViewport({ scale: scaleT });

      const cv = document.createElement('canvas');
      cv.width = vp.width; cv.height = vp.height;
      const cctx = cv.getContext('2d');
      cctx.fillStyle = 'white'; cctx.fillRect(0,0,cv.width,cv.height);
      await page.render({ canvasContext: cctx, viewport: vp }).promise;

      const item = document.createElement('div');
      item.className = 'thumb'; item.dataset.page = i;
      item.appendChild(cv);

      const pspan = document.createElement('div');
      pspan.className = 'pnum'; pspan.textContent = `Page ${i}`;
      item.appendChild(pspan);

      item.addEventListener('click', () => {
        flipScrollTarget = 'top';
        window.setPage(i);
      });

      thumbList.appendChild(item);
      thumbCanvases.push(cv);
    }
    markCurrentThumb(pageNum);
  }

  function markCurrentThumb(n) {
    document.querySelectorAll('.thumb').forEach(el => el.classList.remove('cur'));
    const el = document.querySelector(`.thumb[data-page="${n}"]`);
    if (el) el.classList.add('cur');
    if (el && el.scrollIntoView) el.scrollIntoView({ block: 'nearest' });
  }

  // ====== Inline rename ======
  function beginEditName(id){
    const f = fields.find(t => t.id === id);
    if (!f) return;

    if (editingId && document.getElementById('nameEditor')) {
      document.getElementById('nameEditor').remove();
      editingId = null;
    }
    const box = document.querySelector(`.box[data-id="${id}"]`);
    if (!box) return;

    const cx = box.offsetLeft + box.offsetWidth / 2;
    const cy = box.offsetTop  + box.offsetHeight / 2;
    const w  = Math.max(60, Math.min(260, box.offsetWidth - 12));

    const input = document.createElement('input');
    input.id = 'nameEditor';
    input.type = 'text';
    input.value = f.name;
    input.style.position = 'absolute';
    input.style.left = cx + 'px';
    input.style.top  = cy + 'px';
    input.style.transform = 'translate(-50%,-50%)';
    input.style.width  = w + 'px';
    input.style.padding = '4px 8px';
    input.style.fontFamily = 'system-ui,Segoe UI,Roboto,Arial';
    input.style.fontSize   = '12px';
    input.style.background = '#fff';
    input.style.border = '1px solid #1976d2';
    input.style.borderRadius = '8px';
    input.style.boxShadow = '0 2px 8px rgba(0,0,0,.12)';
    input.style.zIndex = 2147483647;

    overlay.appendChild(input);
    editingId = id;
    input.focus(); input.select();

    function commit(){
      if (!document.getElementById('nameEditor')) return;
      const newName = input.value.trim();
      input.remove(); editingId = null;
      if (newName && newName !== f.name){
        post({ type:'renameField', id:f.id, name:newName, page: pageNum });
      } else {
        redraw();
      }
    }
    function cancel(){
      if (!document.getElementById('nameEditor')) return;
      input.remove(); editingId = null; redraw();
    }

    input.addEventListener('keydown', e => {
      if (e.key === 'Enter') { e.preventDefault(); commit(); }
      if (e.key === 'Escape') { e.preventDefault(); cancel(); }
      e.stopPropagation();
    });
    input.addEventListener('blur', () => commit());
  }

  // ====== Redraw overlay boxes ======
  function redraw(){
    overlay.innerHTML = '';
    if (!viewport) return;
    const Hpt = pageHeightPt();

    fields.forEach(f => {
      const div = document.createElement('div');
      div.className = 'box ' + (f.required ? 'req' : 'nreq') + (f.id===selectedId ? ' sel':'');
      const xpx = ptToPx(f.x), wpx = ptToPx(f.w), hpx = ptToPx(f.h);
      const ypx = ptToPx(Hpt - f.y - f.h);

      div.style.left   = xpx + 'px';
      div.style.top    = ypx + 'px';
      div.style.width  = wpx + 'px';
      div.style.height = hpx + 'px';
      div.title = f.name;
      div.dataset.id = f.id;

      // label giữa box
      const lbl = document.createElement('span');
      lbl.textContent = f.name;
      lbl.style.position   = 'absolute';
      lbl.style.left       = '50%';
      lbl.style.top        = '50%';
      lbl.style.transform  = 'translate(-50%,-50%)';
      lbl.style.padding    = '2px 6px';
      lbl.style.fontFamily = 'system-ui,Segoe UI,Roboto,Arial';
      lbl.style.fontSize   = '12px';
      lbl.style.background = 'rgba(255,255,255,.7)';
      lbl.style.borderRadius = '6px';
      lbl.style.pointerEvents = 'none';
      lbl.style.maxWidth    = '100%';
      lbl.style.whiteSpace  = 'nowrap';
      lbl.style.overflow    = 'hidden';
      lbl.style.textOverflow= 'ellipsis';
      div.appendChild(lbl);

      // dblclick vào box để edit tên
      div.addEventListener('dblclick', (ev) => {
        ev.preventDefault(); ev.stopPropagation();
        beginEditName(f.id);
      });

      if (f.id === selectedId){
        const h = document.createElement('div');
        h.className = 'handle';
        h.dataset.id = f.id;
        div.appendChild(h);
      }
      overlay.appendChild(div);
    });
  }

  // ====== Mouse draw/move/resize (snap + min size) ======
  overlay.addEventListener('mousedown', e => {
    if (editingId) return; // đang edit -> bỏ thao tác chuột

    const rect = overlay.getBoundingClientRect();
    const x = e.clientX - rect.left, y = e.clientY - rect.top;

    const t = e.target;
    if (t.classList.contains('handle')) {
      selectedId = t.dataset.id;
      const f = fields.find(i => i.id===selectedId); if (!f) return;
      dragging = {mode:'resize', sx:x, sy:y, startRect:{...f}};
      overlay.focus();
      e.preventDefault(); return;
    }
    const box = t.closest && t.closest('.box');
    if (box) {
      selectedId = box.dataset.id;
      const f = fields.find(i => i.id===selectedId);
      dragging = {mode:'move', sx:x, sy:y, startRect:{...f}};
      redraw();
      overlay.focus();
      e.preventDefault(); return;
    }

    // draw new
    dragging = {mode:'draw', sx:x, sy:y, startRect:null, moved:false};
    const draft = document.createElement('div');
    draft.className = 'box req';
    draft.style.left = x + 'px'; draft.style.top = y + 'px';
    draft.style.width = '0px'; draft.style.height = '0px';
    draft.id = 'draft';
    overlay.appendChild(draft);
    overlay.focus();
  });

  overlay.addEventListener('mousemove', e => {
    if (!dragging) return;
    const rect = overlay.getBoundingClientRect();
    let x = e.clientX - rect.left, y = e.clientY - rect.top;

    if (dragging.mode === 'draw') {
      const draft = document.getElementById('draft');
      let left = Math.min(dragging.sx, x);
      let top  = Math.min(dragging.sy, y);
      let w = Math.abs(x - dragging.sx);
      let h = Math.abs(y - dragging.sy);

      dragging.moved ||= (w > 2 || h > 2);
      const s = snapRectPx(left, top, w, h);
      draft.style.left = s.left + 'px';
      draft.style.top  = s.top  + 'px';
      draft.style.width = s.width + 'px';
      draft.style.height= s.height + 'px';
      return;
    }

    const f0 = dragging.startRect;
    if (dragging.mode === 'move') {
      const dx = x - dragging.sx, dy = y - dragging.sy;
      const nxpx = ptToPx(f0.x) + dx, nypx = ptToPx(pageHeightPt() - f0.y - f0.h) + dy;
      const s = snapRectPx(nxpx, nypx, ptToPx(f0.w), ptToPx(f0.h));
      const f = fields.find(t => t.id === selectedId);
      if (f) { f.x = pxToPt(s.left); f.y = pageHeightPt() - pxToPt(s.top) - f.h; redraw(); }
      return;
    }

    if (dragging.mode === 'resize') {
      const dx = x - dragging.sx, dy = y - dragging.sy;
      let wpx = Math.max(1, ptToPx(f0.w) + dx);
      let hpx = Math.max(1, ptToPx(f0.h) + dy);
      const s = snapRectPx(ptToPx(f0.x), ptToPx(pageHeightPt() - f0.y - f0.h), wpx, hpx);
      const f = fields.find(t => t.id === selectedId);
      if (f) { f.w = pxToPt(s.width); f.h = pxToPt(s.height); redraw(); }
      return;
    }
  });

  window.addEventListener('mouseup', e => {
    if (!dragging) return;

    if (dragging.mode === 'draw') {
      const draft = document.getElementById('draft');
      if (draft) {
        const xpx = parseFloat(draft.style.left);
        const ypx = parseFloat(draft.style.top);
        const wpx = parseFloat(draft.style.width);
        const hpx = parseFloat(draft.style.height);
        const minPx = MIN_GRID_CELLS * gridPx();

        if (!dragging.moved || wpx < minPx || hpx < minPx) {
          draft.remove();
          dragging = null;
          return; // không tạo field
        }

        draft.remove();
        const Hpt = pageHeightPt();
        const xpt = pxToPt(xpx), wpt = pxToPt(wpx), hpt = pxToPt(hpx);
        const ypt = Hpt - pxToPt(ypx) - hpt;
        post({ type:'addField', page: pageNum, rect: { x:xpt, y:ypt, w:wpt, h:hpt, required:true } });
      }
      dragging = null;
      return;
    }

    // move / resize -> commit
    const f = fields.find(t => t.id === selectedId);
    if (f) post({ type:'updateField', id:f.id, page: pageNum, rect: { x:f.x, y:f.y, w:f.w, h:f.h } });
    dragging = null;
  });

  // ====== Keyboard (flip page + nudge + toggle + delete + inline rename) ======
  overlay.addEventListener('keydown', e => {
    if (editingId) return; // đang edit -> bỏ phím tắt

    // Flip trang bằng phím
    if (e.key === 'PageDown' && pageNum < pdfDoc.numPages) {
      flipScrollTarget = 'top';
      window.setPage(pageNum + 1);
      e.preventDefault();
      return;
    }
    if (e.key === 'PageUp' && pageNum > 1) {
      flipScrollTarget = 'bottom';
      window.setPage(pageNum - 1);
      e.preventDefault();
      return;
    }

    // Field ops
    const f = fields.find(t => t.id === selectedId);
    if (!f) return;
    const stepPt = e.shiftKey ? 5 : 1;

    if (['ArrowLeft','ArrowRight','ArrowUp','ArrowDown'].includes(e.key)) {
      if (e.key==='ArrowLeft')  f.x -= stepPt;
      if (e.key==='ArrowRight') f.x += stepPt;
      if (e.key==='ArrowUp')    f.y += stepPt;
      if (e.key==='ArrowDown')  f.y -= stepPt;
      redraw();
      post({ type:'updateField', id:f.id, page: pageNum, rect: { x:f.x, y:f.y, w:f.w, h:f.h } });
      e.preventDefault();
    }
    if (e.key.toLowerCase()==='r') { post({ type:'toggleRequired', id:f.id, page: pageNum }); e.preventDefault(); }
    if (e.key==='Delete') { post({ type:'deleteField', id:f.id, page: pageNum }); e.preventDefault(); }
    if (e.key==='Enter' || e.key==='F2') { beginEditName(f.id); e.preventDefault(); }
  });

  // Fallback dblclick (phòng dội event)
  overlay.addEventListener('dblclick', (e) => {
    const box = e.target.closest && e.target.closest('.box');
    if (!box) return;
    const id = box.dataset.id;
    if (!id) return;
    e.preventDefault(); e.stopPropagation();
    beginEditName(id);
  });

  // ====== Template bar (drag & drop) ======
  let templatesUi = [];
  function renderTplBar(){
    const list = document.getElementById('tpllist');
    list.innerHTML = '';
    templatesUi.forEach(t => {
      const d = document.createElement('div'); d.className='tpl'; d.draggable=true;
      d.textContent = t.name;
      d.addEventListener('dragstart', ev => {
        ev.dataTransfer.setData('application/json', JSON.stringify(t));
      });
      list.appendChild(d);
    });
  }
  overlay.addEventListener('dragover', e => e.preventDefault());
  overlay.addEventListener('drop', e => {
    e.preventDefault();
    const data = e.dataTransfer.getData('application/json');
    if (!data) return;
    const tpl = JSON.parse(data);
    const rect = overlay.getBoundingClientRect();
    const x = e.clientX - rect.left, y = e.clientY - rect.top;
    const Hpt = pageHeightPt();
    for (const it of tpl.items) {
      const s = snapRectPx(x + ptToPx(it.dx), y + ptToPx(it.dy), ptToPx(it.w), ptToPx(it.h));
      const xpt = pxToPt(s.left);
      const ypt = Hpt - pxToPt(s.top) - it.h;
      post({ type:'addField', page: pageNum, name: it.name, rect: { x:xpt, y:ypt, w:it.w, h:it.h, required: it.required } });
    }
  });

  // ====== Page flip by mouse wheel ======
  function handleWheel(e){
    if (!pdfDoc || editingId || dragging) return;
    const now = Date.now();
    if (now - lastFlipTs < 200) return;               // debounce 200ms

    const canScroll = (container.scrollHeight - container.clientHeight) > 2;
    const nearTop = !canScroll || container.scrollTop <= 0;
    const nearBottom = !canScroll || (container.scrollTop + container.clientHeight >= container.scrollHeight - 2);

    if (e.deltaY > 0 && nearBottom && pageNum < pdfDoc.numPages) {
      e.preventDefault();
      lastFlipTs = now;
      flipScrollTarget = 'top';
      window.setPage(pageNum + 1);
    } else if (e.deltaY < 0 && nearTop && pageNum > 1) {
      e.preventDefault();
      lastFlipTs = now;
      flipScrollTarget = 'bottom';
      window.setPage(pageNum - 1);
    }
  }
  container.addEventListener('wheel', handleWheel, { passive: false });
  overlay.addEventListener('wheel', handleWheel, { passive: false });

  // ====== API host → webview ======
  window.setPage      = async function(num){ pageNum = Math.min(Math.max(1, num), pdfDoc.numPages); await renderPage(pageNum); };
  window.setFields    = function(list){ fields = list || []; if (selectedId && !fields.some(f=>f.id===selectedId)) selectedId=null; redraw(); overlay.focus(); };
  window.zoomIn       = async function(){ scale = Math.min(scale + 0.25, 4); await renderPage(pageNum); };
  window.zoomOut      = async function(){ scale = Math.max(scale - 0.25, 0.5); await renderPage(pageNum); };
  window.setTemplates = function(list){ templatesUi = list || []; renderTplBar(); };

  // ====== Boot ======
  async function main(){
    pdfDoc = await pdfjsLib.getDocument(pdfUrl).promise;
    await buildThumbs();
    await renderPage(pageNum);
  }
  main();
</script>
</body>
</html>

""";

        private string BuildPdfHtml(string pdfFileUri)
        {
            return HtmlTemplate.Replace("__PDF_URL__", pdfFileUri);
        }
    }
}

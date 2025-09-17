using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using static iText.StyledXmlParser.Jsoup.Select.Evaluator;

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
        Label info = new() { AutoSize = true, ForeColor = Color.DimGray };

        public MainForm()
        {
            InitializeComponent(); // dùng Designer thì giữ, không dùng thì có thể xoá
            SetupUi();
        }

        private void SetupUi()
        {
            Text = "PdfSignerStudio (Word Interop + WebView2 + iText7)";
            ClientSize = new Size(1200, 800);
            StartPosition = FormStartPosition.CenterScreen;

            topBar.Controls.AddRange(new Control[] { btnOpenDocx, pageBox, btnZoomOut, btnZoomIn, btnSaveJson, btnLoadJson, btnExport, info });
            btnOpenDocx.Left = 8;
            pageBox.Left = 120;
            btnZoomOut.Left = 210; btnZoomOut.Width = 30;
            btnZoomIn.Left = 245; btnZoomIn.Width = 30;
            btnSaveJson.Left = 290;
            btnLoadJson.Left = 390;
            btnExport.Left = 490;
            info.Left = 590; info.Top = 12;

            Controls.Add(web);
            Controls.Add(topBar);

            btnOpenDocx.Click += OnOpenDocx;
            pageBox.SelectedIndexChanged += (_, __) => SyncPageToWeb();
            btnSaveJson.Click += (_, __) => SaveJson();
            btnLoadJson.Click += (_, __) => LoadJson();
            btnExport.Click += (_, __) => ExportPdf();
            btnZoomIn.Click += async (_, __) => await web.CoreWebView2?.ExecuteScriptAsync("zoomIn()")!;
            btnZoomOut.Click += async (_, __) => await web.CoreWebView2?.ExecuteScriptAsync("zoomOut()")!;
        }

        // Chạy Interop Word trong STA
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

        async void OnOpenDocx(object? sender, EventArgs e)
        {
            using var ofd = new OpenFileDialog { Filter = "Word (*.docx)|*.docx" };
            if (ofd.ShowDialog() != DialogResult.OK) return;

            state = new ProjectState { SourceDocx = ofd.FileName };
            string outDir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "PdfSignerStudio");
            Directory.CreateDirectory(outDir);

            info.Text = "Converting DOCX → PDF with Microsoft Word...";
            try
            {
                state.TempPdf = await RunSTA(() =>
                    PdfService.ConvertDocxToPdfWithWord(ofd.FileName, outDir));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Convert failed: " + ex.Message);
                info.Text = "Convert failed.";
                return;
            }

            info.Text = "Loading preview...";
            await EnsureWebReady();

            web.CoreWebView2.WebMessageReceived -= WebMessageReceived;
            web.CoreWebView2.WebMessageReceived += WebMessageReceived;

            // Map thư mục PDF -> https://app/...
            var pdfFolder = System.IO.Path.GetDirectoryName(state.TempPdf!)!;
            web.CoreWebView2.SetVirtualHostNameToFolderMapping(
                "app", pdfFolder, CoreWebView2HostResourceAccessKind.Allow);

            var pdfUri = $"https://app/{System.IO.Path.GetFileName(state.TempPdf!)}";
            var html = BuildPdfHtml(pdfUri);
            web.CoreWebView2.NavigateToString(html);

            info.Text = "Ready. Drag / move / resize boxes; Double-click rename; R to toggle required; Delete to remove; Zoom +/-.";
        }

        // nhận message từ HTML (pdf.js)
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
                            int numPages = root.GetProperty("numPages").GetInt32();
                            int page = root.GetProperty("page").GetInt32();

                            if (pageBox.Items.Count != numPages)
                            {
                                pageBox.Items.Clear();
                                for (int i = 1; i <= numPages; i++) pageBox.Items.Add($"Page {i}");
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

                            var name = $"Signature_{state.Fields.Count(f => f.Type == "signature") + 1}";
                            var def = new FormFieldDef(name, "signature", page, new RectFpt(x, y, w, h), true);
                            state.Fields.Add(def);

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

                            // tránh trùng tên ngay trong UI
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
            catch
            {
                // ignore malformed messages
            }
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

            // đảm bảo mọi field đều có Id (case load JSON cũ)
            foreach (var f in state.Fields.Where(f => string.IsNullOrEmpty(f.Id)))
            {
                var idx = state.Fields.IndexOf(f);
                state.Fields[idx] = f with { };
            }

            var fields = state.Fields.Where(f => f.Page == page).Select(f => new
            {
                id = f.Id,
                name = f.Name,
                x = f.Rect.X,
                y = f.Rect.Y,
                w = f.Rect.W,
                h = f.Rect.H,
                required = f.Required
            });
            string json = JsonSerializer.Serialize(fields);
            await web.CoreWebView2.ExecuteScriptAsync($"setFields({json});");
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

            state = JsonSerializer.Deserialize<ProjectState>(File.ReadAllText(ofd.FileName)) ?? new ProjectState();
            // gán Id cho field cũ chưa có
            for (int i = 0; i < state.Fields.Count; i++)
            {
                if (string.IsNullOrEmpty(state.Fields[i].Id))
                    state.Fields[i] = state.Fields[i] with { };
            }

            info.Text = "Loaded JSON.";

            if (!string.IsNullOrEmpty(state.TempPdf) && File.Exists(state.TempPdf))
            {
                await EnsureWebReady();

                web.CoreWebView2.WebMessageReceived -= WebMessageReceived;
                web.CoreWebView2.WebMessageReceived += WebMessageReceived;

                var pdfFolder = System.IO.Path.GetDirectoryName(state.TempPdf!)!;
                web.CoreWebView2.SetVirtualHostNameToFolderMapping(
                    "app", pdfFolder, CoreWebView2HostResourceAccessKind.Allow);

                var pdfUri = $"https://app/{System.IO.Path.GetFileName(state.TempPdf!)}";
                var html = BuildPdfHtml(pdfUri);
                web.CoreWebView2.NavigateToString(html);
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

        // HTML: pdf.js + overlay + select/move/resize + zoom
        private string BuildPdfHtml(string pdfFileUri)
        {
            return $@"
<!DOCTYPE html><html>
<head>
<meta charset='utf-8'/>
<meta http-equiv='X-UA-Compatible' content='IE=edge'/>
<meta name='viewport' content='width=device-width, initial-scale=1'/>
<style>
  html, body {{ margin:0; padding:0; height:100%; overflow:hidden; background:#f5f5f5; user-select:none; }}
  #container {{ position:relative; width:100%; height:100%; overflow:auto; }}
  #pagewrap {{ position:relative; margin:24px auto; width:fit-content; }}
  #pdfCanvas {{ display:block; background:white; box-shadow:0 0 20px rgba(0,0,0,.08); }}
  #overlay {{ position:absolute; left:0; top:0; pointer-events:auto; }}
  .box {{ position:absolute; border:2px solid red; box-sizing:border-box; }}
  .box.req {{ border-color:red; }}
  .box.nreq {{ border-color:#888; border-style:dashed; }}
  .box.sel {{ outline:2px solid #1976d2; }}
  .handle {{ position:absolute; width:10px; height:10px; background:#1976d2; right:-6px; bottom:-6px; cursor:nwse-resize; }}
</style>
</head>
<body>
<div id='container'>
  <div id='pagewrap'>
    <canvas id='pdfCanvas'></canvas>
    <div id='overlay' tabindex='0'></div>
  </div>
</div>

<script src='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js'></script>
<script>
  const pdfjsLib = window['pdfjs-dist/build/pdf'];
  pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

  const pdfUrl = '{pdfFileUri}';
  let pdfDoc = null, pageNum = 1, scale = 1.25;
  let viewport = null;
  const canvas = document.getElementById('pdfCanvas');
  const ctx = canvas.getContext('2d');
  const overlay = document.getElementById('overlay');

  let fields = [];        // [{{id,name,x,y,w,h,required}}]
  let selectedId = null;   // id đang chọn
  let dragging = null;    // {{mode:'move'|'resize'|'draw', sx,sy, startRect}}


  function post(msg) {{ chrome.webview.postMessage(JSON.stringify(msg)); }}

  function ptToPx(v) {{ return v * scale; }}
  function pxToPt(v) {{ return v / scale; }}
  function pageHeightPt() {{ return viewport.height / scale; }}

  async function renderPage(num) {{
    const page = await pdfDoc.getPage(num);
    viewport = page.getViewport({{ scale }});
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    overlay.style.width = canvas.width + 'px';
    overlay.style.height = canvas.height + 'px';
    overlay.style.left = '0px';
    overlay.style.top = '0px';
    await page.render({{ canvasContext: ctx, viewport }}).promise;
    post({{ type:'meta', numPages: pdfDoc.numPages, page:num }});
    redraw();
  }}

  function redraw() {{
    overlay.innerHTML = '';
    if (!viewport) return;
    const Hpt = pageHeightPt();
    fields.forEach(f => {{
      const div = document.createElement('div');
      div.className = 'box ' + (f.required ? 'req' : 'nreq') + (f.id===selectedId ? ' sel':'');
      const xpx = ptToPx(f.x);
      const wpx = ptToPx(f.w);
      const hpx = ptToPx(f.h);
      const ypx = ptToPx(Hpt - f.y - f.h);
      div.style.left = xpx + 'px';
      div.style.top = ypx + 'px';
      div.style.width = wpx + 'px';
      div.style.height = hpx + 'px';
      div.title = f.name;
      div.dataset.id = f.id;

      if (f.id === selectedId) {{
        const h = document.createElement('div');
        h.className = 'handle';
        h.dataset.id = f.id;
        div.appendChild(h);
      }}
      overlay.appendChild(div);
    }});
  }}

  // ====== Events ======
  overlay.addEventListener('mousedown', e => {{
    const rect = overlay.getBoundingClientRect();
    const x = e.clientX - rect.left, y = e.clientY - rect.top;

    const target = e.target;
    if (target.classList.contains('handle')) {{
      selectedId = target.dataset.id;
      const f = fields.find(t => t.id === selectedId);
      if (!f) return;
      dragging = {{ mode:'resize', sx:x, sy:y, startRect:{{...f}} }};
      e.preventDefault(); return;
    }}

    const box = target.closest('.box');
    if (box) {{
      selectedId = box.dataset.id;
      const f = fields.find(t => t.id === selectedId);
      dragging = {{ mode:'move', sx:x, sy:y, startRect:{{...f}} }};
      redraw();
      e.preventDefault(); return;
    }}

    // click nền -> bắt đầu vẽ box mới
    dragging = {{ mode:'draw', sx:x, sy:y, startRect:null }};
    const draft = document.createElement('div');
    draft.className = 'box req';
    draft.style.left = x + 'px';
    draft.style.top = y + 'px';
    draft.style.width = '0px';
    draft.style.height = '0px';
    draft.id = 'draft';
    overlay.appendChild(draft);
  }});

  overlay.addEventListener('mousemove', e => {{
    if (!dragging) return;
    const rect = overlay.getBoundingClientRect();
    const x = e.clientX - rect.left, y = e.clientY - rect.top;

    if (dragging.mode === 'draw') {{
      const draft = document.getElementById('draft');
      const left = Math.min(dragging.sx, x);
      const top  = Math.min(dragging.sy, y);
      const w = Math.abs(x - dragging.sx);
      const h = Math.abs(y - dragging.sy);
      draft.style.left = left + 'px';
      draft.style.top = top + 'px';
      draft.style.width = w + 'px';
      draft.style.height = h + 'px';
      return;
    }}

    const f0 = dragging.startRect;
    const Hpt = pageHeightPt();

    if (dragging.mode === 'move') {{
      const dx = x - dragging.sx, dy = y - dragging.sy;
      // đổi sang pt & tính toạ độ mới
      const nx = f0.x + pxToPt(dx);
      const ny = f0.y + pxToPt(-dy);
      const f = fields.find(t => t.id === selectedId);
      if (f) {{ f.x = nx; f.y = ny; redraw(); }}
      return;
    }}

    if (dragging.mode === 'resize') {{
      const dx = x - dragging.sx, dy = y - dragging.sy;
      const nw = Math.max(1, f0.w + pxToPt(dx));
      const nh = Math.max(1, f0.h + pxToPt(dy));
      const f = fields.find(t => t.id === selectedId);
      if (f) {{ f.w = nw; f.h = nh; redraw(); }}
      return;
    }}
  }});

  window.addEventListener('mouseup', e => {{
    if (!dragging) return;

    if (dragging.mode === 'draw') {{
      const draft = document.getElementById('draft');
      if (draft) {{
        const xpx = parseFloat(draft.style.left);
        const ypx = parseFloat(draft.style.top);
        const wpx = parseFloat(draft.style.width);
        const hpx = parseFloat(draft.style.height);
        draft.remove();

        const Hpt = pageHeightPt();
        const xpt = pxToPt(xpx);
        const wpt = pxToPt(wpx);
        const hpt = pxToPt(hpx);
        const ypt = Hpt - pxToPt(ypx) - hpt;

        post({{ type:'addField', page: pageNum, rect: {{ x:xpt, y:ypt, w:wpt, h:hpt }} }});
      }}
      dragging = null; return;
    }}

    if (dragging.mode === 'move' || dragging.mode === 'resize') {{
      const f = fields.find(t => t.id === selectedId);
      if (f) {{
        post({{ type:'updateField', id:f.id, page: pageNum, rect: {{ x:f.x, y:f.y, w:f.w, h:f.h }} }});
      }}
      dragging = null; return;
    }}
  }});

  // keyboard: Delete (xoá), R (toggle), Enter (rename)
  overlay.addEventListener('keydown', e => {{
    if (!selectedId) return;
    const f = fields.find(t => t.id === selectedId);
    if (!f) return;

    if (e.key === 'Delete') {{
      post({{ type:'deleteField', id:f.id, page: pageNum }});
      e.preventDefault();
    }} else if (e.key.toLowerCase() === 'r') {{
      post({{ type:'toggleRequired', id:f.id, page: pageNum }});
      e.preventDefault();
    }} else if (e.key === 'Enter' || e.key === 'F2') {{
      const name = prompt('New field name:', f.name);
      if (name) post({{ type:'renameField', id:f.id, name, page: pageNum }});
      e.preventDefault();
    }}
  }});

  // double-click để rename
  overlay.addEventListener('dblclick', e => {{
    const box = e.target.closest('.box');
    if (!box) return;
    selectedId = box.dataset.id;
    const f = fields.find(t => t.id === selectedId);
    if (!f) return;
    const name = prompt('New field name:', f.name);
    if (name) post({{ type:'renameField', id:f.id, name, page: pageNum }});
  }});

  // API host → webview
  window.setPage = async function(num) {{
    pageNum = Math.min(Math.max(1, num), pdfDoc.numPages);
    await renderPage(pageNum);
  }};

  window.setFields = function(list) {{
    fields = list || [];
    // nếu đang chọn id đã bị xoá, clear
    if (selectedId && !fields.some(f => f.id === selectedId)) selectedId = null;
    redraw();
    overlay.focus();
  }};

  // Zoom
  window.zoomIn = async function() {{
    scale = Math.min(scale + 0.25, 4);
    await renderPage(pageNum);
  }};
  window.zoomOut = async function() {{
    scale = Math.max(scale - 0.25, 0.5);
    await renderPage(pageNum);
  }};

  async function main() {{
    pdfDoc = await pdfjsLib.getDocument(pdfUrl).promise;
    await renderPage(pageNum);
  }}
  main();
</script>
</body>
</html>";
        }
    }
}

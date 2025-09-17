using System;
using System.Collections.Generic;

namespace PdfSignerStudio;

public record RectFpt(float X, float Y, float W, float H);

// Giữ nguyên chữ ký cũ để không phải sửa chỗ khác; thêm Id mặc định.
public record FormFieldDef(
    string Name,
    string Type,      // "signature"
    int Page,         // 1-based
    RectFpt Rect,
    bool Required = false
)
{
    public string Id { get; init; } = Guid.NewGuid().ToString();
}

public class ProjectState
{
    public string? SourceDocx { get; set; }
    public string? TempPdf { get; set; }
    public float PreviewDpi { get; set; } = 150f;
    public List<FormFieldDef> Fields { get; set; } = new();
}

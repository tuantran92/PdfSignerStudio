using System;
using System.Collections.Generic;

namespace PdfSignerStudio;

public record RectFpt(float X, float Y, float W, float H);

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

// ====== Template library ======
public record TemplateField(
    string Name,
    float W, float H,
    bool Required = true,
    float Dx = 0, float Dy = 0   // offset (pt) từ điểm thả
);

public record TemplateDef(
    string Name,
    List<TemplateField> Items
);

public class ProjectState
{
    public string? SourceDocx { get; set; }
    public string? TempPdf { get; set; }
    public float PreviewDpi { get; set; } = 150f;
    public List<FormFieldDef> Fields { get; set; } = new();
}

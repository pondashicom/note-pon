namespace NotePon;

internal enum NoteTextAlignment
{
    Left,
    Center,
    Right,
    Justify
}

internal sealed record FormattedNoteDocument(
    FormattedNoteParagraph[] Paragraphs);

internal sealed record FormattedNoteParagraph(
    FormattedNoteRun[] Runs,
    NoteTextAlignment Alignment,
    int IndentLevel,
    string? BulletText);

internal sealed record FormattedNoteRun(
    string Text,
    bool Bold,
    bool Italic,
    bool Underline,
    bool Superscript,
    bool Subscript,
    int? RgbColor);

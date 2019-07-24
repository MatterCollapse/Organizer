using System;
using System.ComponentModel;
using System.ComponentModel.Design.Serialization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms.Integration;
using System.Windows.Forms.Design;
using System.Windows.Documents;
using System.IO;
using System.Windows.Media;
using System.Drawing;

[Designer(typeof(ControlDesigner))]
//[DesignerSerializer("System.Windows.Forms.Design.ControlCodeDomSerializer, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", "System.ComponentModel.Design.Serialization.CodeDomSerializer, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")]
public class SpellBox : ElementHost
{
    private int WORD_WRAP_WIDTH = 2000;

    public SpellBox()
    {
        box = new RichTextBox();
        base.Child = box;
        box.TextChanged += (s, e) => OnTextChanged(EventArgs.Empty);
        box.SpellCheck.IsEnabled = true;
        box.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
        this.Size = new System.Drawing.Size(100, 20);

        indent = 0;
    }
    public override string Text
    {
        get {
            return new TextRange(box.Document.ContentStart, box.Document.ContentEnd).Text;
        }
        set {
            box.Document.Blocks.Clear();
            box.Document.Blocks.Add(new Paragraph(new Run(value)));
        }
    }

    internal void Select(TextRange selection, System.Drawing.Color color)
    {
        //TextRange selection = new TextRange(box.Document.ContentStart.GetPositionAtOffset(index), box.Document.ContentStart.GetPositionAtOffset(index + length));
        selection.ApplyPropertyValue(TextElement.FontWeightProperty, FontWeights.Bold);
        selection.ApplyPropertyValue(FlowDocument.ForegroundProperty, new SolidColorBrush(System.Windows.Media.Color.FromArgb(color.A, color.R, color.G, color.B)));
        box.Selection.Select(selection.Start, selection.End);
        box.Focus();
    }

    public bool WordWrap
    {
        get
        {
            return box.Document.PageWidth == WORD_WRAP_WIDTH;
        }
        set
        {
            if (value)
            {
                box.Document.PageWidth = box.Width;
            }
            else
            {
                box.Document.PageWidth = WORD_WRAP_WIDTH;
            }
        }
    }

    TextPointer GetTextPositionAtOffset(TextPointer position, int characterCount)
    {
        while (position != null)
        {
            if (position.GetPointerContext(LogicalDirection.Forward) == TextPointerContext.Text)
            {
                int count = position.GetTextRunLength(LogicalDirection.Forward);
                if (characterCount <= count)
                {
                    return position.GetPositionAtOffset(characterCount);
                }

                characterCount -= count;
            }

            TextPointer nextContextPosition = position.GetNextContextPosition(LogicalDirection.Forward);
            if (nextContextPosition == null)
                return position;

            position = nextContextPosition;
        }

        return position;
    }

    public TextRange FindTextInRange(TextRange searchRange, string searchText)
    {
        int offset = searchRange.Text.IndexOf(searchText, StringComparison.OrdinalIgnoreCase);
        if (offset < 0)
            return null;  // Not found

        var start = GetTextPositionAtOffset(searchRange.Start, offset);
        TextRange result = new TextRange(start, GetTextPositionAtOffset(start, searchText.Length));
        return result;
    }

    public TextRange Find(string text, TextPointer lastFound)
    {
        TextRange searchRange = new TextRange(null != lastFound ? lastFound : box.Document.ContentStart, box.Document.ContentEnd);
        TextRange foundRange = FindTextInRange(searchRange, text);

        return foundRange;
    }

    public float ZoomFactor
    {
        get
        {
            return 1.0f;
        }
        set
        {
            //change ZoomFactor
        }
    }

    public HorizontalAlignment SelectionAlignment
    {
        get
        {
            object val = box.Selection.GetPropertyValue(RichTextBox.HorizontalContentAlignmentProperty);
            return (HorizontalAlignment) box.Selection.GetPropertyValue(RichTextBox.HorizontalContentAlignmentProperty);
        }
        set
        {
            box.Selection.ApplyPropertyValue(RichTextBox.HorizontalContentAlignmentProperty, value);
        }
    }

    [DefaultValue(false)]
    public bool Multiline
    {
        get { return box.AcceptsReturn; }
        set { box.AcceptsReturn = value; }
    }
    
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public new System.Windows.UIElement Child
    {
        get { return base.Child; }
        set { /* Do nothing to solve a problem with the serializer !! */ }
    }

    public bool SelectionBullet {
        get { return true; }
        set
        {
            EditingCommands.ToggleBullets.Execute(null, box);
        }
    }
    public int SelectionIndent {
        get { return indent; }
        set
        {
            indent += value;
            if (box.Selection.Text != string.Empty)
                box.Selection.Text = string.Empty;

            var caretPosition = box.CaretPosition.GetPositionAtOffset(0, LogicalDirection.Forward);

            string spaceText = "";
            for (int i = 0; i < indent; i++)
                spaceText += " ";

            box.CaretPosition.InsertTextInRun(spaceText);
            box.CaretPosition = caretPosition;
        }
    }

    public void Clear()
    {
        box.Document.Blocks.Clear();
        box.Document.Blocks.Add(new Paragraph(new Run("")));
    }

    public void LoadFile(string filePath)
    {
        TextRange range;
        FileStream fStream;
        if (File.Exists(filePath))
        {
            range = new TextRange(box.Document.ContentStart, box.Document.ContentEnd);
            fStream = new FileStream(filePath, FileMode.OpenOrCreate);
            range.Load(fStream, DataFormats.Rtf);
            fStream.Close();
        }
    }

    public void SaveFile(string filePath)
    {
        TextRange range;
        FileStream fStream;
        range = new TextRange(box.Document.ContentStart, box.Document.ContentEnd);
        fStream = new FileStream(filePath, FileMode.Create);
        range.Save(fStream, DataFormats.Rtf);
        fStream.Close();
    }

    private void PrintCommand()
    {
        PrintDialog pd = new PrintDialog();
        if ((pd.ShowDialog() == true))
        {
            //use either one of the below      
            pd.PrintVisual(box as Visual, "printing as visual");
            //pd.PrintDocument((((IDocumentPaginatorSource)box.Document).DocumentPaginator), "printing as paginator");
        }
    }

    public int SelectionLength
    {
        get { return box.Selection.Text.Length; }   
    }

    public System.Drawing.Color SelectionBackColor
    {
        set {
            
            box.Selection.ApplyPropertyValue(FlowDocument.BackgroundProperty, new SolidColorBrush(System.Windows.Media.Color.FromArgb(value.A, value.R, value.G, value.B))); }
    }

    public int SelectionStart
    {
        get { return box.Selection.Start.GetOffsetToPosition(box.Document.ContentStart); }
    }

    private RichTextBox box;
    private int indent;

    internal void Paste()
    {
        box.Paste();
    }

    internal void Cut()
    {
        box.Cut();
    }

    internal void Copy()
    {
        box.Copy();
    }

    internal void SelectAll()
    {
        box.SelectAll();
    }

    internal void Undo()
    {
        box.Undo();
    }

    internal void Redo()
    {
        box.Redo();
    }
}
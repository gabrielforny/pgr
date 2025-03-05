import win32com.client 
from win32com.client import constants

class Find_Replace():
    def __init__(self, word_in: str):
        self.word_in = word_in
        self.word_app = win32com.client.DispatchEx("Word.Application")
        self.word_app.Visible = False
        self.word_app.DisplayAlerts = False
        self.doc = self.word_app.Documents.Open(word_in)
        self.stop_centralizing = False  # Flag para parar de centralizar após certo termo

    def replace_text(self, find_str: str, replace_with: str):
        find = self.word_app.Selection.Find
        find.Text = find_str
        find.Replacement.Text = replace_with
        find.Forward = True
        find.Wrap = 1
        find.Format = True
        find.MatchCase = True
        find.MatchWholeWord = False
        find.MatchWildcards = True

        print(f"🔄 Substituindo: {find_str} → {replace_with}")

        while find.Execute():
            selection = self.word_app.Selection
            selection.Text = replace_with
            
            # Se já encontrou "MÊS DE VIGENCIA...", não centraliza mais
            if not self.stop_centralizing:
                selection.ParagraphFormat.Alignment = 1  # Centraliza
            
            # Verifica se é o gatilho para parar de centralizar
            if find_str == "MÊS DE VIGENCIA ANO VIGENCIA A MÊS DE VIGENCIA ANO VIGENCIA":
                self.stop_centralizing = True

            # Remove destaque
            selection.Range.Shading.BackgroundPatternColor = 0    
            selection.Range.HighlightColorIndex = 0

    def replace_in_paragraphs(self, find_str: str, replace_with: str):
        for para in self.doc.Paragraphs:
            if find_str in para.Range.Text:
                print(f"📝 Substituindo em Parágrafo: {find_str} → {replace_with}")
                para.Range.Text = para.Range.Text.replace(find_str, replace_with)

                if not self.stop_centralizing:
                    para.Range.ParagraphFormat.Alignment = 1  # Centraliza
                
                if find_str == "MÊS DE VIGENCIA ANO VIGENCIA A MÊS DE VIGENCIA ANO VIGENCIA":
                    self.stop_centralizing = True

                para.Range.Font.Shading.BackgroundPatternColor = 0    
                para.Range.Font.HighlightColorIndex = 0

    def replace_in_shapes(self, find_str: str, replace_with: str):
        for shape in self.doc.Shapes:
            if shape.TextFrame.HasText:
                text_range = shape.TextFrame.TextRange
                if find_str in text_range.Text:
                    print(f"🖼️ Substituindo em Shape: {find_str} → {replace_with}")
                    text_range.Text = text_range.Text.replace(find_str, replace_with)

                    if not self.stop_centralizing:
                        text_range.ParagraphFormat.Alignment = 1  # Centraliza
                    
                    if find_str == "MÊS DE VIGENCIA ANO VIGENCIA A MÊS DE VIGENCIA ANO VIGENCIA":
                        self.stop_centralizing = True

                    text_range.Font.Shading.BackgroundPatternColor = 0    
                    text_range.Font.HighlightColorIndex = 0

    def replace_in_headers_and_footers(self, find_str: str, replace_with: str):
        for section in self.doc.Sections:
            for header in section.Headers:
                if header.Exists:
                    header_range = header.Range
                    if find_str in header_range.Text:
                        print(f"🔝 Substituindo no Cabeçalho: {find_str} → {replace_with}")
                        header_range.Text = header_range.Text.replace(find_str, replace_with)

                        if not self.stop_centralizing:
                            header_range.ParagraphFormat.Alignment = 1  # Centraliza
                        
                        if find_str == "MÊS DE VIGENCIA ANO VIGENCIA A MÊS DE VIGENCIA ANO VIGENCIA":
                            self.stop_centralizing = True

                        header_range.Font.Shading.BackgroundPatternColor = 0    
                        header_range.Font.HighlightColorIndex = 0

            for footer in section.Footers:
                if footer.Exists:
                    footer_range = footer.Range
                    if find_str in footer_range.Text:
                        print(f"🔻 Substituindo no Rodapé: {find_str} → {replace_with}")
                        footer_range.Text = footer_range.Text.replace(find_str, replace_with)

                        if not self.stop_centralizing:
                            footer_range.ParagraphFormat.Alignment = 1  # Centraliza
                        
                        if find_str == "MÊS DE VIGENCIA ANO VIGENCIA A MÊS DE VIGENCIA ANO VIGENCIA":
                            self.stop_centralizing = True

                        footer_range.Font.Shading.BackgroundPatternColor = 0    
                        footer_range.Font.HighlightColorIndex = 0

    def save_close_file(self, word_out: str):
        self.doc.SaveAs(word_out)
        self.doc.Close()
        self.word_app.Application.Quit()

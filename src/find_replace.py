import win32com.client 
from win32com.client import constants

class Find_Replace():
    def __init__(self, word_in: str):
        '''
        wd_replace => 2=replace all, 1=replace one 
        wd_find_wrap => 2=ask to continue, 1=continue search
        '''
        self.word_in = word_in
        self.word_app = win32com.client.DispatchEx("Word.Application")
        self.word_app.Visible = False
        self.word_app.DisplayAlerts = False
        # Abrir documento
        self.doc = self.word_app.Documents.Open(word_in)

    def replace_text(self, find_str: str, replace_with: str):
        """Substitui texto no corpo do documento mantendo formatação e removendo destaque"""
        find = self.word_app.Selection.Find
        find.Text = find_str
        find.Replacement.Text = replace_with
        find.Forward = True
        find.Wrap = 1  # wdFindContinue
        find.Format = False
        find.MatchCase = True
        find.MatchWholeWord = False
        find.MatchWildcards = True

        print(f"🔄 Substituindo: {find_str} → {replace_with}")

        while find.Execute():
            selection = self.word_app.Selection
            selection.Text = replace_with
            
            opcoes_centralizadas = [
                "MÊS DE VIGENCIA ANO VIGENCIA A MÊS DE VIGENCIA ANO VIGENCIA",
                "XX.XX.XXXX",
                "NOME DA EMPRESA",
                "varCnae",
                "varGrauRisco",
                "varDescCnae",
                "qtdFunMas",
                "qtdFunFem",
                "qtdFunMenor",
                "qtdFunTotal",
                "varGrauRisc",
                "varCol31",
                "varCol32",
                "varCol41",
                "varCol42",
                "varCol51",
                "varCol52",
                "varCol61",
                "varCol62",
                "varCol71",
                "varCol72",
                "varCol81",
                "varCol82",
                "varCol91",
                "varCol92",
                "varCol101",
                "varCol102",
                "varCol111",
                "varCol112",
                "varCol121",
                "varCol122",
                "varCol131",
                "varCol132",
                "varCol141",
                "varCol142",
                "varCol151",
                "varCol152"
            ]
            if find_str in opcoes_centralizadas:
                selection.ParagraphFormat.Alignment = 1
            elif find_str == "CURITIBA/PR, 00 de Abril de 2024":
                selection.ParagraphFormat.Alignment = 2
            else:
                selection.ParagraphFormat.Alignment = 0
                
            # Remover destaque corretamente
            # selection.Range.Shading.BackgroundPatternColor = 0    
            # selection.Range.HighlightColorIndex = 0

    def replace_in_paragraphs(self, find_str: str, replace_with: str):
        """Substitui texto nos parágrafos mantendo formatação e removendo destaque"""
        for para in self.doc.Paragraphs:
            if find_str in para.Range.Text:
                print(f"📝 Substituindo em Parágrafo: {find_str} → {replace_with}")
                para.Range.Text = para.Range.Text.replace(find_str, replace_with)
                para.Range.ParagraphFormat.Alignment = 1  # Mantém centralizado
                
                # Remover destaque corretamente
                # para.Range.Font.Shading.BackgroundPatternColor = 0    
                # para.Range.Font.HighlightColorIndex = 0

    def replace_in_shapes(self, find_str: str, replace_with: str):
        """Substitui texto dentro de Shapes (Caixas de Texto) mantendo formatação e removendo destaque"""
        for shape in self.doc.Shapes:
            if shape.TextFrame.HasText:
                text_range = shape.TextFrame.TextRange
                if find_str in text_range.Text:
                    print(f"🖼️ Substituindo em Shape: {find_str} → {replace_with}")
                    text_range.Text = text_range.Text.replace(find_str, replace_with)
                    text_range.ParagraphFormat.Alignment = 1  # Mantém centralizado
                    
                    # Remover destaque corretamente
                    # text_range.Font.Shading.BackgroundPatternColor = 0    
                    # text_range.Font.HighlightColorIndex = 0

    def replace_in_headers_and_footers(self, find_str: str, replace_with: str):
        """Substitui texto nos Cabeçalhos e Rodapés mantendo formatação e removendo destaque"""
        for section in self.doc.Sections:
            for header in section.Headers:
                if header.Exists:
                    header_range = header.Range
                    if find_str in header_range.Text:
                        print(f"🔝 Substituindo no Cabeçalho: {find_str} → {replace_with}")
                        header_range.Text = header_range.Text.replace(find_str, replace_with)
                        header_range.ParagraphFormat.Alignment = 1  # Mantém centralizado
                        
                        # Remover destaque corretamente
                        # header_range.Font.Shading.BackgroundPatternColor = 0    
                        # header_range.Font.HighlightColorIndex = 0

            for footer in section.Footers:
                if footer.Exists:
                    footer_range = footer.Range
                    if find_str in footer_range.Text:
                        print(f"🔻 Substituindo no Rodapé: {find_str} → {replace_with}")
                        footer_range.Text = footer_range.Text.replace(find_str, replace_with)
                        footer_range.ParagraphFormat.Alignment = 1  # Mantém centralizado
                        
                        # Remover destaque corretamente
                        # footer_range.Font.Shading.BackgroundPatternColor = 0    
                        # footer_range.Font.HighlightColorIndex = 0
                    
    def save_close_file(self, word_out: str):
        """Salva e fecha o documento"""
        self.doc.SaveAs(word_out)
        self.doc.Close()
        self.word_app.Application.Quit()

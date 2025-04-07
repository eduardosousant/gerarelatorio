import locale
import os
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
from PIL import Image as PILImage
from matplotlib.ticker import MaxNLocator
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.utils import ImageReader
from reportlab.platypus import BaseDocTemplate, Paragraph, Spacer, Table, Frame, PageTemplate, Image, TableStyle, PageBreak, NextPageTemplate
from reportlab.lib.pagesizes import landscape
from reportlab.platypus import NextPageTemplate



def obter_caminho_arquivo(arquivo):
    if getattr(sys, 'frozen', False):
        return os.path.join(sys._MEIPASS, arquivo)
    return os.path.join(os.getcwd(), arquivo)

locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')



def header(canvas, doc, mes_ano, data_iniciobr, data_fimbr):
    canvas.saveState()
    logo_path = 'logo.png'
    styles = getSampleStyleSheet()
    try:
        img = PILImage.open(logo_path)
        width, height = img.size
        aspect_ratio = height / width
        new_width = 110
        new_height = new_width * aspect_ratio
        img = Image(logo_path, width=new_width, height=new_height)
    except:
        img = ''
    else:
        pass
    data = [[img, Paragraph('<b>RELATÓRIO DE CONSUMO</b><br/>' + '<b>Cliente:</b> Ambipar Environment<br/>' + f'<b>Período:</b> {data_iniciobr} a {data_fimbr}', styles['Normal'])]]
    table = Table(data, colWidths=[120, A4[0] - 160])
    table.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'MIDDLE'), ('LEFTPADDING', (1, 0), (1, 0), 10)]))
    w, h = table.wrap(A4[0] - 80, 100)
    table.drawOn(canvas, 40, A4[1] - h - 10)
    canvas.line(40, A4[1] - h - 10, A4[0] - 40, A4[1] - h - 10)
    canvas.restoreState()

def footer(canvas, doc):
    canvas.saveState()
    texto_rodape = 'assistentecomercial01@geoeste.com.br'
    canvas.setFont('Helvetica', 7)
    text_width = canvas.stringWidth(texto_rodape, 'Helvetica', 8)
    canvas.drawString((A4[0] - text_width) / 2, 30, texto_rodape)
    canvas.setFont('Helvetica', 8)
    canvas.drawRightString(A4[0] - 40, 30, f'Página {doc.page}')
    canvas.restoreState()

def gerar_grafico(consumo_diario, mes_ano, data_inicio, data_fim):
    largura_polegadas = 4.91
    altura_polegadas = 3
    max_consumo = consumo_diario['Hidrômetro'].max()
    plt.figure(figsize=(largura_polegadas, altura_polegadas))
    plt.plot(consumo_diario['data'], consumo_diario['Hidrômetro'], marker='o', linestyle='-', color='teal', markersize=2, linewidth=0.5, alpha=0.8)
    plt.title(f'Consumo Diário de Água no periodo')
    plt.xlabel('Data', fontsize=4, family='sans-serif')
    plt.ylabel('Consumo (m³)', fontsize=4, family='sans-serif')
    plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%d/%m'))
    plt.gca().xaxis.set_major_locator(mdates.DayLocator(interval=2))
    plt.xticks(rotation=45, ha='right', fontsize=4)
    plt.ylim(0, max_consumo + 25)
    plt.yticks(fontsize=4)
    plt.gca().yaxis.set_major_locator(MaxNLocator(integer=True, prune='lower'))
    plt.grid(True, linestyle='--', color='gray', alpha=0.7)
    plt.tight_layout()
    plt.savefig('grafico_consumo.png', dpi=300)
    plt.close()

def gerar_relatorio(nome_arquivo, mes_ano, consumo_diario, total_consumo, maior_consumo, media_mensal, consumo_semanal_summary, data_inicio, data_fim):
    doc = BaseDocTemplate(nome_arquivo, pagesize=A4)
    styles = getSampleStyleSheet()

    # Convertendo datas para o formato brasileiro
    data_iniciobr = data_inicio.strftime('%d/%m/%Y')
    data_fimbr = data_fim.strftime('%d/%m/%Y')

    # Corrigir chamada do header (passando datas formatadas)
    frame_retrato = Frame(30, 30, A4[0] - 60, A4[1] - 80)
    template_retrato = PageTemplate(id='retrato', frames=[frame_retrato], onPage=lambda c, d: header(c, d, mes_ano, data_iniciobr, data_fimbr), onPageEnd=footer)

    frame_paisagem = Frame(30, 30, A4[1] - 60, A4[0] - 80)
    template_paisagem = PageTemplate(id='paisagem', frames=[frame_paisagem], pagesize=landscape(A4), onPage=lambda c, d: header(c, d, mes_ano, data_iniciobr, data_fimbr), onPageEnd=footer)

    doc.addPageTemplates([template_retrato, template_paisagem])

    # Criando a lista de elementos do relatório
    elements = []

    texto_explicativo = ('O presente relatório apresenta a análise do consumo diário de água, com base nas leituras registradas e organizadas cronologicamente.'
                         ' A partir dessa análise, é possível identificar os períodos de maior demanda e subsidiar a tomada de decisões estratégicas para a otimização do uso da água. '
                         'Essa abordagem visa aprimorar a eficiência operacional e reforçar o compromisso da empresa com a sustentabilidade e a gestão responsável dos recursos hídricos.')

    estilo = styles['Normal']
    estilo.firstLineIndent = 35
    estilo.alignment = 0

    elements.append(Spacer(1, 12))
    elements.append(Paragraph(texto_explicativo, estilo))
    elements.append(Spacer(1, 12))

    # Informações de consumo
    media_mensal_consumo = media_mensal['Hidrômetro'].mean()
    consumo_dias_semana = consumo_semanal_summary[consumo_semanal_summary['final_de_semana'] == 'Dia de Semana'][
        'Hidrômetro'].sum()
    consumo_finais_semana = consumo_semanal_summary[consumo_semanal_summary['final_de_semana'] == 'Final de Semana'][
        'Hidrômetro'].sum()
    elements = [Spacer(1, 12), Paragraph(texto_explicativo, estilo), Spacer(1, 12),
                Paragraph(        f"O total de consumo para o período <b>{data_iniciobr}</b> a <b>{data_fimbr}</b>"
                                  f" foi de <b>{locale.format_string('%.2f', total_consumo, grouping=True)} m³</b>."
                                  f" O maior consumo diário foi de <b>{locale.format_string('%.2f', maior_consumo['Hidrômetro'], grouping=True)} m³</b>."
                                  f" A média mensal foi de <b>{locale.format_string('%.2f', media_mensal_consumo, grouping=True)} m³</b>."
                                  f" O consumo de água foi registrado para os diferentes tipos de consumo, separando os dados de final de semana e dias úteis."
                                  f"  Os valores de consumo total (em m³) para cada tipo são os seguintes: " + 'enquanto que no   '.join(
            [f"{tipo}: <b>{locale.format_string('%.2f', consumo, grouping=True)} m³</b>" for tipo, consumo in
             zip(consumo_semanal_summary['final_de_semana'], consumo_semanal_summary['Hidrômetro'])]) + '.',
        styles['Normal'])]
    elements.append(Spacer(1, 12))

    # Tabela de consumo diário
    tabela_dados = [['Data', 'Consumo (m³)']] + [[data.strftime('%d/%m/%Y'), locale.format_string('%.2f', consumo, grouping=True)] for data, consumo in zip(consumo_diario['data'], consumo_diario['Hidrômetro'])]
    tabela = Table(tabela_dados, style=[
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('ALIGN', (1, 1), (-1, -1), 'RIGHT'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black)
    ])
    elements.append(tabela)

    elements.append(Spacer(1, 12))

    # Adiciona quebra de página antes do gráfico e muda para o template paisagem
    elements.append(NextPageTemplate('paisagem'))
    elements.append(PageBreak())

    # Página paisagem para o gráfico
    elements.append(Paragraph('Gráfico de Consumo Diário', styles['Title']))
    largura_desejada = 750
    image_width, image_height = ImageReader('grafico_consumo.png').getSize()
    proporcao = largura_desejada / image_width
    altura_desejada = image_height * proporcao
    elements.append(Image('grafico_consumo.png', width=largura_desejada, height=altura_desejada))

    # Se houver mais conteúdo depois do gráfico, volte para retrato
    elements.append(NextPageTemplate('retrato'))

    # Construção do documento
    doc.build(elements)

    os.remove('grafico_consumo.png')
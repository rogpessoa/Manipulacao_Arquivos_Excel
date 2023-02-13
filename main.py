"""
 #Banda Inferior
    #formula = MEDIA MOVEL(20PERIODOS) - 2X DESVIO_PADRAO(20P)


    #Banda Superior
    # formula = MEDIA MOVEL(20PERIODOS) + 2X DESVIO_PADRAO(20P)
"""
from datetime import date
from openpyxl.chart import LineChart, Reference
from openpyxl.drawing.image import Image
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.workbook import Workbook
from classes import LeitorAcoes, GerenciadorPlanilha, PropriedadeSerieGrafico

try:
    # acao = input('Qual ação você deseja processar? ').upper()
    acao = 'BIDI4'
    leitor_acoes = LeitorAcoes(caminho_arquivo='./dados/')
    leitor_acoes.processa_arquivo(acao)

    # Criando planilha excel

    gerenciador = GerenciadorPlanilha()
    planilha_dados = gerenciador.adiciona_planilha('Dados')
    gerenciador.adiciona_linha(['DATA', 'COTAÇÃO', 'BANDA INFERIOR', 'BANDA SUPERIOR'])

    indice = 2
    for linha in leitor_acoes.dados:
        # Data
        ano_mes_dia = linha[0].split(" ")[
            0]  # quebra a linha 1 em duas pegando apenas o valor da linha zero >> 2018-05-10
        data = date(
            year=int(ano_mes_dia.split('-')[0]),
            month=int(ano_mes_dia.split('-')[1]),
            day=int(ano_mes_dia.split('-')[2])
        )
        # Cotacao
        cotacao = float(linha[1])

        formula_bb_inferior = f'=AVERAGE(B{indice}:B{indice + 19}) - 2*STDEV(B{indice}:B{indice + 19})'
        formula_bb_superior = f'=AVERAGE(B{indice}:B{indice + 19}) + 2*STDEV(B{indice}:B{indice + 19})'

        gerenciador.atualiza_celula(celula=f'A{indice}', dado=data)
        gerenciador.atualiza_celula(celula=f'B{indice}', dado=cotacao)
        gerenciador.atualiza_celula(celula=f'C{indice}', dado=formula_bb_inferior)
        gerenciador.atualiza_celula(celula=f'D{indice}', dado=formula_bb_superior)

        indice += 1
    gerenciador.adiciona_planilha(titulo_planilha='Grafico')

    # Mesclagem das celulas
    gerenciador.mescla_celulas(celula_inicio='A1', celula_fim='T2')

    # Criando os estilos da Planilha

    gerenciador.aplica_estilos(
        celula='A1',
        estilos=[('font', Font(bold=True, size=18, color='FFFFFF')),
                 ('alignment', Alignment(vertical='center', horizontal='center')),
                 ('fill', PatternFill('solid', fgColor='24b8bf'))

                 ]
    )

    referencia_cotacoes = Reference(planilha_dados, min_col=2, min_row=2, max_col=4, max_row=indice)
    referencia_datas = Reference(planilha_dados, min_col=1, min_row=2, max_col=1, max_row=indice)

    gerenciador.atualiza_celula('A1', 'Histórico de Cotações')
    gerenciador.adiciona_grafico_linha(
        celula='A3',
        comprimento=33.87,
        altura=14.82,
        titulo=f'Cotações - {acao}',
        titulo_eixo_x='Data da Cotação',
        titulo_eixo_y='Valor da Cotação',
        referencia_eixo_x=referencia_cotacoes,
        referencia_eixo_y=referencia_datas,
        propriedades_grafico=[
            PropriedadeSerieGrafico(grossura=0, cor_preenchimento='125ce6'),
            PropriedadeSerieGrafico(grossura=0, cor_preenchimento='ba1818'),
            PropriedadeSerieGrafico(grossura=0, cor_preenchimento='05f278')
        ]
    )

    gerenciador.mescla_celulas(celula_inicio='I32', celula_fim='L35')
    gerenciador.adiciona_imagem(celula='I32', caminho_imagem='./recursos/logo.png')
    gerenciador.salva_arquivo('./saida/PlanilhaRefatorada.xlsx')

except FileNotFoundError:
    print('Arquivo não encontrado!')

except ValueError:
    print('Formato de dados incorreto, favor verificar!')

except AttributeError:
    print('Atributo inexistente!')

except Exception as excecao:
    print(f'Ocorreu um erro Inesperado! Erro: \033[0;31m{str(excecao)}\033[m')

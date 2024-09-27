import os
import tabula
import re
import pandas as pd

def extrair_dados_pdfs(diretorio_pdf):
    # Expressões regulares para identificar padrões de telefone e candidato a prefeito
    padrao_telefone = re.compile(r'\d{9}')
    padrao_candidato = re.compile(r'Prefeito\s+\d+\s+([A-ZÁÉÍÓÚÂÊÔÃÕÇ]+(?: [A-ZÁÉÍÓÚÂÊÔÃÕÇ]+)+)')

    # Lista para armazenar os resultados finais
    dados = []

    # Lê cada arquivo PDF no diretório
    for arquivo in os.listdir(diretorio_pdf):
        if arquivo.endswith(".pdf"):
            caminho_pdf = os.path.join(diretorio_pdf, arquivo)
            print(f"Lendo arquivo: {arquivo}")
            
            try:
                # Extrai as tabelas de todas as páginas do PDF
                lista_tabelas = tabula.read_pdf(caminho_pdf, pages="all", multiple_tables=True)

                # Resetar variáveis para garantir que não haja dados persistentes de arquivos anteriores
                ddd = None
                telefone_atendente = None
                tipo_contato = None
                nome_atendente = None
                candidato_prefeito = None
                telefone_prefeito = None

                # "Pesquisa" sobre cada tabela e busca pelas informações
                for tabela in lista_tabelas:
                    try:
                        # Verifica se a tabela tem pelo menos 4 colunas para evitar erros
                        if len(tabela.columns) >= 4:
                            primeira_linha = tabela.iloc[0]
                            
                            # Pega os valores relevantes com verificações para NaN
                            ddd = str(int(primeira_linha[0])).strip() if pd.notna(primeira_linha[0]) else "Não disponível"
                            telefone_atendente = str(int(primeira_linha[1])).strip() if pd.notna(primeira_linha[1]) else "Nenhum telefone encontrado"
                            tipo_contato = str(primeira_linha[2]).strip() if pd.notna(primeira_linha[2]) else "Não disponível"
                            nome_atendente = str(primeira_linha[3]).strip() if pd.notna(primeira_linha[3]) else "Nenhum atendente encontrado"

                            # Formatação correta do telefone e DDD
                            if re.match(r'\d{2}', ddd) and re.match(r'\d{9}', telefone_atendente):
                                telefone_atendente = f"({ddd}) {telefone_atendente[:5]}-{telefone_atendente[5:]}"
                            
                        # Converte a tabela para string para buscar candidatos e telefones
                        tabela_str = tabela.to_string()
                        candidatos_encontrados = padrao_candidato.findall(tabela_str)
                        telefones_encontrados = padrao_telefone.findall(tabela_str)
                        
                        # Verifica se encontrou o candidato a prefeito
                        if candidatos_encontrados and not candidato_prefeito:
                            candidato_prefeito = candidatos_encontrados[0]

                        # Verifica se encontrou o telefone do candidato
                        if telefones_encontrados and not telefone_prefeito:
                            telefone_prefeito = telefones_encontrados[0]

                    except Exception as e:
                        print(f"Erro ao processar a tabela em {arquivo}: {e}")

                # Certifica-se de que todos os campos estão preenchidos
                if not nome_atendente:
                    nome_atendente = 'Nenhum atendente encontrado'
                if not ddd:
                    ddd = 'Não disponível'
                if not telefone_atendente:
                    telefone_atendente = 'Nenhum telefone encontrado'
                if not tipo_contato:
                    tipo_contato = 'Não disponível'
                if not candidato_prefeito:
                    candidato_prefeito = 'Nenhum candidato encontrado'
                if not telefone_prefeito:
                    telefone_prefeito = 'Nenhum telefone encontrado'

                # Adiciona o resultado à lista de dados
                dados.append({
                    'Arquivo': arquivo,
                    'Atendente': nome_atendente,
                    'DDD': ddd,
                    'Telefone Atendente': telefone_atendente,
                    'Tipo': tipo_contato,
                    'Candidato a Prefeito': candidato_prefeito,
                    'Telefone Prefeito': telefone_prefeito
                })

            except Exception as e:
                print(f"Erro ao processar o arquivo {arquivo}: {e}")
    
    # Cria um DataFrame com os dados extraídos
    df = pd.DataFrame(dados, columns=[
        'Arquivo', 
        'Atendente', 
        'DDD', 
        'Telefone Atendente', 
        'Tipo', 
        'Candidato a Prefeito', 
        'Telefone Prefeito'
    ])
    
    # Salva o DataFrame em um arquivo Excel
    df.to_excel("Candidatos_Prefeitos_Atendentes_Telefones.xlsx", index=False)
    print("Planilha Excel criada com sucesso!")

# Defina o caminho da pasta com PDFs abaixo, nesse primeiro diretorio_pdf comentado é referente a pasta 
# DrapTeste desse mesmo arquivo, que é um exemplo do que esse script faz, para conseguir fazer esse teste é
# Importante mudar o "Usuário" abaixo para o referente a sua máquina, deve-se também descomentar a linha abaixo e 
# Comentar o outro diretorio_pdf

#diretorio_pdf = r"C:\Users\Usuário\Desktop\Ler Pasta com PDF Python\DrapTeste"

diretorio_pdf = r"G:\Meu Drive\Drap"
# Essa linha acima refere-se a uma pasta que deve estar dentro do Drive, onde o app Drive deve estar instalado na 
# Máquina e com uma pasta com o nome Drap e com arquivos PDFS como no Exemplo.pdf e como os arquivos dentro da pasta DrapTeste

# Chama a função para ler todos os PDFs na pasta e salvar os resultados em uma planilha
extrair_dados_pdfs(diretorio_pdf)

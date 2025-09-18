import re
import string
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

from wordcloud import WordCloud
from collections import Counter


def limpar_texto(texto):
    if pd.isna(texto):
        return ""
    
    # Converter para string e minúsculas
    texto = str(texto).lower()
    
    # Remover pontuações e números
    texto = re.sub(r'[^\w\s]', ' ', texto)
    texto = re.sub(r'\d+', '', texto)
    
    # Remover espaços extras
    texto = ' '.join(texto.split())
    
    return texto


def remover_stopwords(texto, stopwords_personalizadas=None):
    """
    Remove stopwords comuns em português
    """
    stopwords_pt = {
        'a','ao','aos','aquela','aquelas','aquele','aqueles','aquilo','as','até','com','como',
        'da','das','de','dela','delas','dele','deles','depois','do','dos','e','ela','elas',
        'ele','eles','em','entre','essa','essas','esse','esses','esta','está','estamos',
        'estão','estar','estas','estava','estavam','este','esteja','estejam','estejamos',
        'estes','esteve','estive','estivemos','estiver','estivera','estiveram','estiverem',
        'estivermos','estivesse','estivessem','estivéssemos','estou','eu','foi','fomos',
        'for','fora','foram','forem','formos','fosse','fossem','fôssemos','fui','há',
        'haja','hajam','hajamos','hão','havemos','haver','hei','houve','houvemos',
        'houver','houvera','houveram','houverei','houverem','houveremos','houveria',
        'houveriam','houveríamos','houvermos','houvesse','houvessem','houvéssemos',
        'isso','isto','já','lhe','lhes','mais','mas','me','mesmo','meu','meus',
        'minha','minhas','muito','na','não','nas','nem','no','nos','nós','nossa',
        'nossas','nosso','nossos','num','numa','o','os','ou','para','pela','pelas',
        'pelo','pelos','por','qual','quando','que','quem','se','seja','sejam',
        'sejamos','sem','ser','será','serão','serem','seremos','seria','seriam',
        'seríamos','seu','seus','só','somos','sou','sua','suas','são','também',
        'te','tem','temos','tenha','tenham','tenhamos','tenho','ter','terá','terão',
        'terem','teremos','teria','teriam','teríamos','teve','tinha','tinham',
        'tínhamos','tive','tivemos','tiver','tivera','tiveram','tiverem','tivermos',
        'tivesse','tivessem','tivéssemos','tu','tua','tuas','tudo','um','uma','você',
        'vocês','vos','à','às','é','fiz','pra','pos','pós','preciso','pois','dia', 
        'solicitou','porque','realizado','causa','onde','fazer','meses','após','vezes', 
        'ano','dias','indicada','encaminhada','então','paciente','uso','anos','desde',
        'observamos','cerca','nao','data','encaminho'
    }
    
    if stopwords_personalizadas:
        stopwords_pt.update(stopwords_personalizadas)
    
    palavras = texto.split()
    palavras_filtradas = [palavra for palavra in palavras if palavra not in stopwords_pt and len(palavra) > 2]
    
    return ' '.join(palavras_filtradas)


def verificar_arquivo_csv(arquivo_planilha):
    print("=== VERIFICANDO ARQUIVO CSV ===")
    
    try:
        with open(arquivo_planilha, 'r', encoding='utf-8') as f:
            linhas = f.readlines()[:10]
        
        print(f"Primeiras linhas do arquivo:")
        for i, linha in enumerate(linhas, 1):
            print(f"Linha {i}: {repr(linha)}")
        
        delimitadores = [',', ';', '\t', '|']
        
        for delim in delimitadores:
            try:
                df_teste = pd.read_csv(arquivo_planilha, delimiter=delim, nrows=5)
                print(f"\n✅ Sucesso com delimitador '{delim}'")
                print(f"Colunas encontradas: {list(df_teste.columns)}")
                print(f"Shape: {df_teste.shape}")
                return delim
            except Exception as e:
                print(f"❌ Falhou com delimitador '{delim}': {str(e)[:100]}")
        
        return None
        
    except Exception as e:
        print(f"Erro ao verificar arquivo: {e}")
        return None


def gerar_tabela_frequencias(texto_final, nome_arquivo_saida="frequencias.csv"):
    """
    Gera tabela com frequência absoluta, relativa e acumulada
    das 10 principais palavras e salva em planilha.
    """
    palavras = texto_final.split()
    contador = Counter(palavras)

    total_palavras = sum(contador.values())
    top10 = contador.most_common(10)

    df_freq = pd.DataFrame(top10, columns=["Palavra", "Frequência Absoluta"])
    df_freq["Frequência Relativa"] = df_freq["Frequência Absoluta"] / total_palavras
    df_freq["Frequência Acumulada"] = df_freq["Frequência Relativa"].cumsum()

    df_freq.to_csv(nome_arquivo_saida, index=False)
    print(f"\n✅ Tabela de frequências salva como: {nome_arquivo_saida}")

    return df_freq


def criar_nuvem_palavras(arquivo_planilha, nome_coluna, titulo, stopwords_extra=None, 
                        salvar_imagem=True, nome_arquivo_saida="nuvem_palavras.png",
                        nome_arquivo_freq="frequencias.csv"):
    try:
        if arquivo_planilha.endswith('.csv'):
            delimitador = verificar_arquivo_csv(arquivo_planilha)
            
            if delimitador is None:
                print("\n=== TENTANDO LEITURA ROBUSTA ===")
                try:
                    df = pd.read_csv(arquivo_planilha, 
                                   encoding='utf-8',
                                   sep=None,
                                   engine='python',
                                   quotechar='"',
                                   skipinitialspace=True,
                                   on_bad_lines='skip')
                except:
                    df = pd.read_csv(arquivo_planilha, 
                                   encoding='utf-8',
                                   sep='\n',
                                   header=None,
                                   names=['texto_completo'])
            else:
                df = pd.read_csv(arquivo_planilha, delimiter=delimitador, encoding='utf-8')
        else:
            df = pd.read_excel(arquivo_planilha)
        
        print(f"\n✅ Planilha carregada com sucesso! {len(df)} linhas encontradas.")
        print(f"Colunas disponíveis: {list(df.columns)}")
        
        if nome_coluna not in df.columns:
            print(f"❌ Erro: Coluna '{nome_coluna}' não encontrada!")
            print(f"Colunas disponíveis: {list(df.columns)}")
            return None
        
        textos = df[nome_coluna].fillna('')
        
        print(f"\n=== AMOSTRA DOS DADOS ===")
        for i, texto in enumerate(textos.head(3)):
            print(f"Linha {i+1}: {str(texto)[:100]}...")
        
        texto_completo = ' '.join(textos.astype(str))
        print(f"\nTexto combinado tem {len(texto_completo)} caracteres.")
        
        texto_limpo = limpar_texto(texto_completo)
        texto_final = remover_stopwords(texto_limpo, stopwords_extra)
        
        if not texto_final.strip():
            print("❌ Erro: Nenhum texto válido encontrado após processamento!")
            return None
        
        # 🔹 Geração da tabela de frequências
        tabela = gerar_tabela_frequencias(texto_final, nome_arquivo_saida=nome_arquivo_freq)
        print("\nTop 10 palavras com frequências:")
        print(tabela)
        
        # Configurar nuvem de palavras
        wordcloud = WordCloud(
            width=1200,
            height=600,
            background_color='white',
            max_words=100,
            colormap='viridis',
            relative_scaling=0.5,
            min_font_size=10
        ).generate(texto_final)
        
        plt.figure(figsize=(15, 8))
        plt.imshow(wordcloud, interpolation='bilinear')
        plt.axis('off')
        plt.title(titulo, fontsize=24, pad=20)
        
        if salvar_imagem:
            plt.savefig(nome_arquivo_saida, dpi=300, bbox_inches='tight')
            print(f"\n✅ Nuvem de palavras salva como: {nome_arquivo_saida}")
        
        plt.show()
        
        return wordcloud
    
    except Exception as e:
        print(f"❌ Erro ao processar: {str(e)}")
        import traceback
        traceback.print_exc()
        return None


# Exemplo de uso
if __name__ == "__main__":
    arquivo_prontuario = "coluna_dados_paciente_limpo.csv" 
    coluna_prontuario = "Table 1"
    
    nuvem = criar_nuvem_palavras(
        arquivo_planilha=arquivo_prontuario,
        nome_coluna=coluna_prontuario,
        titulo="Prontuários",
        salvar_imagem=True,
        nome_arquivo_saida="paciente_nuvem.png",
        nome_arquivo_freq="frequencias_prontuario.csv"
    )

    arquivo_queixa = "teste_coluna_limpo.csv"
    coluna_queixa = "texto"

    nuvem_prontuario = criar_nuvem_palavras(
        arquivo_planilha=arquivo_queixa,
        nome_coluna=coluna_queixa,
        titulo="Queixas",
        salvar_imagem=True,
        nome_arquivo_saida="prontuario_nuvem.png",
        nome_arquivo_freq="frequencias_queixas.csv"
    )

import streamlit as st
import pandas as pd
import io

# Define a classe ContaBancaria
class ContaBancaria:
    def __init__(self, empresa, banco, agencia, conta):
        self.empresa = empresa
        self.banco = banco
        self.agencia = agencia
        self.conta = conta

    def exibir_detalhes(self):
        return f"Empresa: {self.empresa}, Banco: {self.banco}, Agência: {self.agencia}, Conta: {self.conta}"

# Instâncias
conta_corporeos = ContaBancaria("Corporeos", "341", "0285", "682977")
conta_elfran = ContaBancaria("Elfran", "033", "1042", "130005033")
conta_alisa = ContaBancaria("Alisa", "033", "3413", "130031533")

# Funções para selecionar conta
def selecionar_banco(filial):
    return {
        '0101': conta_corporeos.banco,
        '6401': conta_elfran.banco,
        '7901': conta_alisa.banco
    }.get(filial, 'Banco Desconhecido')

def selecionar_agencia(filial):
    return {
        '0101': conta_corporeos.agencia,
        '6401': conta_elfran.agencia,
        '7901': conta_alisa.agencia
    }.get(filial, 'Agencia Desconhecida')

def selecionar_conta(filial):
    return {
        '0101': conta_corporeos.conta,
        '6401': conta_elfran.conta,
        '7901': conta_alisa.conta
    }.get(filial, 'Conta Desconhecida')

# Título do App
st.title("📊 Rateio Sodexo - Alimentação e Refeição")

# Upload do arquivo
uploaded_file = st.file_uploader("📎 Envie o arquivo Excel", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Carrega as planilhas
        df_corporeos = pd.read_excel(uploaded_file, sheet_name=1)
        df_ELFRAN = pd.read_excel(uploaded_file, sheet_name=2)
        df_ALISA = pd.read_excel(uploaded_file, sheet_name=3)

        # Função para tratar os dados
        def processar_df(df, filial):
            df = df.reset_index(drop=True)
            df.columns = df.iloc[0]
            df = df[1:].reset_index(drop=True)
            df['Filial'] = filial
            return df

        df_corporeos = processar_df(df_corporeos, '0101')
        df_ELFRAN = processar_df(df_ELFRAN, '6401')
        df_ALISA = processar_df(df_ALISA, '7901')

        # Concatenação
        df_concatenado = pd.concat([df_corporeos, df_ELFRAN, df_ALISA], ignore_index=True)

        # Renomear colunas específicas
        df_concatenado.columns.values[1] = 'Conta Contabil Refeição'
        df_concatenado.columns.values[2] = 'Valor Refeição'
        df_concatenado.columns.values[3] = 'Conta Contabil Alimentação'
        df_concatenado.columns.values[4] = 'Valor Alimentação'

        # Filtrar linhas
        df_filtrado = df_concatenado[~df_concatenado['CENTRO DE CUSTO'].str.startswith('TOTAL:')]

        # Separar Alimentação e Refeição
        df_ali = df_filtrado[['Filial', 'Conta Contabil Alimentação', 'Valor Alimentação', 'CENTRO DE CUSTO']].copy()
        df_ref = df_filtrado[['Filial', 'Conta Contabil Refeição', 'Valor Refeição', 'CENTRO DE CUSTO']].copy()

        # Enriquecimento dos dados
        for df, tipo, natureza, historico, valor_col in [
            (df_ali, 'R', '202513', 'Credito Sodexo Alimentacao', 'Valor Alimentação'),
            (df_ref, 'R', '200251', 'Credito Sodexo Refeição', 'Valor Refeição')
        ]:
            df['Numerario'] = 'M1'
            df['Tipo'] = tipo
            df['Valor'] = df[valor_col]
            df['Natureza'] = natureza
            df['Banco'] = df['Filial'].apply(selecionar_banco)
            df['Agencia'] = df['Filial'].apply(selecionar_agencia)
            df['Conta'] = df['Filial'].apply(selecionar_conta)
            df['Num Cheque'] = ''
            df['Historico'] = historico
            df['C. Custo Debito'] = ''
            df['C. Custo Credito'] = df['CENTRO DE CUSTO']
            df['Item debito'] = ''
            df['Item Credito'] = ''
            df['Cl Val D'] = ''
            df['Cl Val C'] = ''
            df['Data'] = 'Inserir data'

        colunas_finais = ['Filial', 'Data', 'Numerario', 'Tipo', 'Valor', 'Natureza',
                          'Banco', 'Agencia', 'Conta', 'Num Cheque', 'Historico',
                          'C. Custo Debito', 'C. Custo Credito', 'Item debito',
                          'Item Credito', 'Cl Val D', 'Cl Val C']

        df_ref_final = df_ref[colunas_finais]
        df_ali_final = df_ali[colunas_finais]

        # Mostrar tabelas
        st.subheader("📄 Visualização - Alimentação")
        st.dataframe(df_ali_final)

        st.subheader("📄 Visualização - Refeição")
        st.dataframe(df_ref_final)

        # Converter para Excel e oferecer download
        def gerar_excel(df, nome_arquivo):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False)
            return output.getvalue()

        st.download_button("⬇️ Baixar Rateio Alimentação", gerar_excel(df_ali_final, 'RateioAlimentacao.xlsx'), file_name='RateioAlimentacao.xlsx')
        st.download_button("⬇️ Baixar Rateio Refeição", gerar_excel(df_ref_final, 'RateioRefeicao.xlsx'), file_name='RateioRefeicao.xlsx')

    except Exception as e:
        st.error(f"❌ Erro ao processar: {e}")

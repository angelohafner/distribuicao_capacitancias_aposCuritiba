import streamlit as st
from funcoes_dupla_estrela import *
from funcoes_streamlit import *
from desenho_dos_capacitores_v2 import CapacitoresY

st.markdown("#### Entre com o número de capacitores em série, em paralelo e com o número de bancos")
# Parâmetros da matriz de capacitores
col1, col2, col3 = st.columns(3)
with col1:
    n_lin = st.number_input("Número de capacitores série",     min_value=0, max_value=10, step=1, format='%d', value=2)
with col2:
    n_col = st.number_input("Número de capacitores paralelos", min_value=0, max_value=10, step=1, format='%d', value=2)
with col3:
    n_ban = st.number_input("Número de bancos de capacitores", min_value=0, max_value=10, step=1, format='%d', value=1)

st.markdown(f"Configuração modelo de um banco de capacitores dupla estrela isolada (se repete para cada banco de capacitor)")
fig = CapacitoresY.generate_and_save_capacitor_plot(m=n_lin, n=n_col, d=2.0, horizontal_spacing=5, filename='capacitores.png')
st.pyplot(fig)


# cada matriz corresponde a um ramo
# um sistema trifásico tem no mínimo 3 ramos
# um sistema trifásico dupla estrela tem no mínimo 6 ramos
capacitancias_por_matriz = n_col * n_lin
num_matrizes = 6 * n_ban
num_capacitancias = num_matrizes * capacitancias_por_matriz

st.markdown("#### Instruções")
st.markdown(f"Uma matriz representa a forma como os capacitores estão ligadas em um ramo, aqui  = {capacitancias_por_matriz}")
st.markdown(f"O número de matrizes ou ramos é o produto do número de fases (3) por 2 (dupla estrela) \
            pelo número de banco de capacitores ({n_ban}) = {num_matrizes}")
st.markdown(f"O número total de capacitores é o produto do número de ramos ({num_matrizes}) pelo número\
            de capacitores por ramo ({capacitancias_por_matriz})= {num_capacitancias}")
st.markdown("Faça o donwload da planilha modelo, onde deverá preencher o valor dos capacitores (coluna B) e o correspondente\
            número de série (coluna A).")
st.markdown("Depois faça o upload da planilha salva para que o programa execute a distribuição dos capacitores\
            em forma de matriz de forma que a associação resulte no conjunto mais equilibrado possível.")
st.markdown("Por fim faça o donwload dos arquivos com a melhor configuração. Um deles mostra cada valor de \
            capacitância e outro o número de série correspondente a cada posição.")

st.markdown("##### Planilha de entrada")
# Ler capacitâncias e número de série correspondente
# Botão de download para o arquivo minhas_latas.xlsx
with open("minhas_latas.xlsx", "rb") as file:
    st.download_button(
        label="Download da planilha modelo",
        data=file,
        file_name="minhas_latas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
# Widget de upload de arquivo
uploaded_file = st.file_uploader("Faça o upload do seu arquivo aqui:", type=['xlsx'])

if uploaded_file is not None:
    with open("minhas_latas.xlsx", "wb") as file:
        file.write(uploaded_file.getbuffer())
    st.success("Arquivo carregado e salvo como minhas_latas.xlsx")


# lendo apenas as primeiras num_capacitancias da tabela pra nao precisar ficando trocando a tabela de entrada
capacitancias_df = pd.read_excel("minhas_latas.xlsx").head(num_capacitancias)
capacitancias = np.array(capacitancias_df["uF"])
num_serie_cap = np.array(capacitancias_df["no_serie"], dtype=str)

# Criar as matrizes de capacitores e matrizes número de série
matrizes, numeros = criar_matrizes(num_capacitancias,
                                    capacitancias, num_serie_cap,
                                    capacitancias_por_matriz, n_lin, n_col)

# Encontrar melhor configuração
n_iteracoes = 10000
melhor_configuracao_value, melhor_configuracao_nrser = \
    otimizar_capacitancias(n_iteracoes, matrizes, numeros, n_lin, n_col,
                           calcular_capacitancia_equivalente)





exportar_matrizes_para_excel(melhor_configuracao_nrser, melhor_configuracao_value,
                             'melhor_configuracao_nrser.xlsx', 'melhor_configuracao_value.xlsx')


st.markdown("##### Planilha de saída")
file_paths = ["melhor_configuracao_nrser.xlsx", "melhor_configuracao_value.xlsx"]
zip_name = "configuracoes.zip"
create_zip_file(file_paths, zip_name)
# Botão de download para o arquivo zip
with open(zip_name, "rb") as fp:
    st.download_button(
        label="Download do melhor arranjo de capacitores",
        data=fp,
        file_name=zip_name,
        mime="application/zip"
    )
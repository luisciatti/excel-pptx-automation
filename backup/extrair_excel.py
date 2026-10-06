import pandas as pd
import unicodedata

def normalizar_nome(nome):
    if pd.isna(nome):
        return ""
    nome = str(nome).strip().lower()
    nome = unicodedata.normalize("NFKD", nome).encode("ASCII", "ignore").decode("utf-8")
    return nome

# Dicionário de equivalências entre nomes diferentes
mapa_clientes = {
    "fundacion teleton vida": "hospital fundacion teleton",
    "idf - instituto doutor feitosa": "hospital idf (instituto doutor feitosa)",
    "hnso - hospital nossa senhora de oliveira": "hns oliveira",
    "hob - hospital oftalmologico de brasilia - grupo patria": "grupo opty",
    "neotin - sao francisco hospital & maternidade": "neotin neonatal terapia intensiva ltda",
    "proton - centro medico campinas": "proton servicos radiologicos ltda",
    "hospital santa emilia": "hospital e maternidade santa emilia",
    "onco clinica jaragua ltda": "onco jaragua",
    "furj univille - fundacao educacional da regiao de joinville": "univille (faculdade)",
    "tecsalud - fundacion santos y de la garza evia, i.b.p": "tecsalud",
    "the american british cowdray medical center i.a.p.": "cmabc",
    "inc - instituto de neurocirurgia capixaba ltda": "instituto de neurocirurgia capixaba ltda inc",
    "pio sodalicio das damas de caridade de caxias do sul": "pompeia",
    "fundacao hospitalar santa terezinha de erechim": "erechim"
}

def gerar_relatorio(storage_file, incidentes_file, ano, output_file):
    # --- BLOCO 1: Relatório base ---
    df = pd.read_csv(storage_file, header=[0,1], encoding="latin1")

    colunas = []
    mes_atual = ""
    for mes, tipo in df.columns:
        if "Unnamed" not in mes:
            mes_atual = mes
        colunas.append((mes_atual, tipo))
    df.columns = pd.MultiIndex.from_tuples(colunas)

    resultado = []
    for _, row in df.iterrows():
        cliente = row.xs("Customer", level=1).values[0]
        meses = set([col[0] for col in df.columns])
        for mes in meses:
            usuarios = None
            s3 = None
            database = None
            for m, tipo in df.columns:
                if m == mes:
                    if "Users" in tipo:
                        usuarios = row[(m, tipo)]
                    if "S3" in tipo:
                        s3 = row[(m, tipo)]
                    if "Database" in tipo:
                        database = row[(m, tipo)]
            if usuarios is not None and s3 is not None and database is not None:
                if not pd.isna(usuarios) and not pd.isna(s3) and not pd.isna(database):
                    resultado.append({
                        "Cliente": cliente,
                        "Mês": mes,
                        "Usuários": usuarios,
                        "S3 (Gb)": s3,
                        "Database (Gb)": database
                    })

    df_final = pd.DataFrame(resultado)
    df_final["Mês"] = pd.to_datetime(df_final["Mês"])
    df_final["Disponibilidade"] = df_final["Mês"].dt.days_in_month * 24
    df_final = df_final.sort_values(["Cliente", "Mês"])
    df_final["Mês"] = df_final["Mês"].dt.date

    # --- BLOCO 2: Incidentes ---
    df_inc = pd.read_csv(incidentes_file, encoding="latin1", skiprows=4)
    df_inc = df_inc.loc[:, ~df_inc.columns.str.contains("Unnamed")]

    meses = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho",
             "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"]

    dados = []
    for _, row in df_inc.iterrows():
        cliente = row["Customers"]
        for i, mes in enumerate(meses):
            base = i*5+1
            horas = row.iloc[base+1]
            minutos = row.iloc[base+2]
            nro_inc = row.iloc[base+3]
            avail = row.iloc[base+4]
            dados.append({
                "Cliente": cliente,
                "Mês": pd.to_datetime(f"{ano}-{i+1:02d}-01").date(),
                "Horas": horas,
                "Minutos Indisponibilidade": minutos,
                "Incidentes": nro_inc,
                "Disponibilidade (%)": avail
            })

    df_inc_clean = pd.DataFrame(dados)

    # --- Normalizar e mapear nomes ---
    df_final["Cliente"] = df_final["Cliente"].apply(normalizar_nome)
    df_inc_clean["Cliente"] = df_inc_clean["Cliente"].apply(normalizar_nome)
    df_inc_clean["Cliente"] = df_inc_clean["Cliente"].replace(mapa_clientes)

    # --- Checar diferenças ---
    clientes_storage = set(df_final["Cliente"].unique())
    clientes_incidentes = set(df_inc_clean["Cliente"].unique())
    print("Clientes sem incidentes:", clientes_storage - clientes_incidentes)
    print("Clientes sem storage:", clientes_incidentes - clientes_storage)

    # --- BLOCO 3: Merge ---
    df_final = pd.merge(
        df_final,
        df_inc_clean,
        on=["Cliente","Mês"],
        how="left"
    )

    df_final.to_excel(output_file, index=False)
    print(f"Relatório {ano} criado: {output_file}")


# Gerar relatórios para 2025 e 2026
gerar_relatorio(
    r"C:\Users\320311014\Códigos\Excel\Excel_project\Excel\Storage 2025.csv",
    r"C:\Users\320311014\Códigos\Excel\Excel_project\Excel\Incidentes 2025.csv",
    2025,
    "Relatorio2025_Com_Incidentes9.xlsx"
)

gerar_relatorio(
    r"C:\Users\320311014\Códigos\Excel\Excel_project\Excel\Storage 2026.csv",
    r"C:\Users\320311014\Códigos\Excel\Excel_project\Excel\Incidentes 2026.csv",
    2026,
    "Relatorio2026_Com_Incidentes9.xlsx"
)

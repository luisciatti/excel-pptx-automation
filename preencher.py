# =========================
# 📂 IMPORTS
# =========================
import sys
import re
import unicodedata
import calendar
import copy
import os
import shutil
import subprocess
import tempfile
from collections import Counter
from pathlib import Path

import pandas as pd
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN

import paths

CLIENT_MAP = {
    "einstein": "hospital israelita albert einstein",
    "hospital idf": "hospital idf (instituto doutor feitosa)",
    "casas andre luiz": "centro espirita nosso lar - casas andre luiz",
    "spplus": "spdm - associacao paulista para o desenvolvimento da medicina",
    "inc": "instituto de neurocirurgia capixaba ltda inc",
    "neotin": "neotin neonatal terapia intensiva ltda",
    "proton": "proton servicos radiologicos ltda",
    "cmabc": "the american british cowdray medical center i.a.p.",
    "hospital espirita marilia": "hospital espirita marilia",
    "instituto atena de pesquisa clinica": "instituto atena de pesquisa clinica ltda",
    "onco jaragua": "onco jaragua",
    "univille": "univille (faculdade)",
    "iot": "iot joinville",
    "iot joinville": "iot joinville",
    "santahelena": "clinica santa helena",
    "clinica santa helena": "clinica santa helena",
    "hob": "opty",
    "grupo opty": "opty",
    "hns": "hns oliveira",
    "hns oliveira": "hns oliveira",
    "hospital das nacoes": "hospital das nacoes",
    "hospital e maternidade santa emilia": "hospital e maternidade santa emilia",
    "hospital fundacion teleton": "hospital fundacion teleton",
    "hospital mexico americano": "hospital mexico americano",
    "mexico americano": "hospital mexico americano",
    "mexicoamericano": "hospital mexico americano",
    "hospital santa emilia": "hospital e maternidade santa emilia",
    "santaemilia": "hospital e maternidade santa emilia",
    "onco clinica jaragua ltda": "onco jaragua",
    "tecsalud - fundacion santos y de la garza evia, i.b.p": "tecsalud",
    "servicios de salud de yucatan": "servicios de salud de yucatan",
    "teleton": "hospital fundacion teleton",
    "yucatan": "servicios de salud de yucatan",
    "hmabc": "hmabc",
    "anestblu": "anestblu",
    "pio sodalicio das damas de caridade de caxias do sul": "pompeia",
    "pompeia": "pompeia",
    "erechim": "erechim",
}

MONTH_NAME_TO_NUMBER = {
    "janeiro": 1,
    "fevereiro": 2,
    "marco": 3,
    "abril": 4,
    "maio": 5,
    "junho": 6,
    "julho": 7,
    "agosto": 8,
    "setembro": 9,
    "outubro": 10,
    "novembro": 11,
    "dezembro": 12,
    # Spanish month names
    "enero": 1,
    "febrero": 2,
    "marzo": 3,
    "mayo": 5,
    "junio": 6,
    "julio": 7,
    "septiembre": 9,
    "octubre": 10,
    "noviembre": 11,
    "diciembre": 12,
}

MONTH_NUMBER_TO_NAME_PT = {
    1: "janeiro",
    2: "fevereiro",
    3: "março",
    4: "abril",
    5: "maio",
    6: "junho",
    7: "julho",
    8: "agosto",
    9: "setembro",
    10: "outubro",
    11: "novembro",
    12: "dezembro",
}

MONTH_NUMBER_TO_NAME_ES = {
    1: "enero",
    2: "febrero",
    3: "marzo",
    4: "abril",
    5: "mayo",
    6: "junio",
    7: "julio",
    8: "agosto",
    9: "septiembre",
    10: "octubre",
    11: "noviembre",
    12: "diciembre",
}

def normalizar(texto):
    if texto is None or pd.isna(texto):
        return ""
    texto = str(texto).strip().lower()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", texto)


def ensure_columns(df, required_columns, source_name):
    missing = [column for column in required_columns if column not in df.columns]
    if missing:
        raise ValueError(f"{source_name} sem colunas obrigatórias: {', '.join(missing)}")


def is_report_outdated(report_path, source_paths):
    if not report_path.exists():
        return True
    try:
        report_mtime = report_path.stat().st_mtime
    except OSError:
        return True
    for source_path in source_paths:
        try:
            if source_path.exists() and source_path.stat().st_mtime > report_mtime:
                return True
        except OSError:
            return True
    return False


def extrair_cliente_mes(texto):
    if texto is None:
        return None, None
    texto = str(texto).replace("\n", " ").replace("\r", " ").replace("\x0b", " ")
    texto = normalizar(texto)
    cliente_match = re.search(r"cliente:\s*(.*?)\s*mes:", texto, re.IGNORECASE)
    mes_match = re.search(r"mes:\s*([a-zç]+[-–]\d{4})", texto, re.IGNORECASE)
    cliente = cliente_match.group(1).strip() if cliente_match else None
    mes = mes_match.group(1).replace("–", "-").strip() if mes_match else None
    return cliente, mes


def normalize_template_stem(stem):
    return re.sub(r"[\s\-_]+", "-", normalizar(stem)).strip("-")


def is_spanish_template_stem(stem):
    return bool(re.search(r"[-–]\s*(mx|mex)$", stem, re.IGNORECASE))


def build_codigo_mes(ano, mes, suffix=None):
    codigo = f"{ano}{mes.zfill(2)}"
    if suffix:
        suffix = suffix.strip().lstrip("-").upper()
        if suffix in ("MX", "MEX"):
            return f"{codigo}-{suffix}"
    return codigo


def build_target_month_label(codigo_mes, language="pt"):
    base_code, _ = split_codigo_mes(codigo_mes)
    ano = base_code[:4]
    mes_num = int(base_code[4:6])
    month_name_map = MONTH_NUMBER_TO_NAME_ES if language == "es" else MONTH_NUMBER_TO_NAME_PT
    mes_nome = month_name_map.get(mes_num)
    if not mes_nome:
        raise ValueError(f"Mês inválido para código {codigo_mes}")
    return f"{mes_nome}-{ano}"


def replace_month_label(text, target_month_label):
    if not text:
        return text
    month_pattern = r"(?:janeiro|fevereiro|marco|março|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro|enero|febrero|marzo|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)"
    return re.sub(rf"{month_pattern}\s*[-–]\s*\d{{4}}", target_month_label, text, flags=re.IGNORECASE)


def duplicate_template_with_month(source_path, target_path, codigo_mes, language="pt"):
    prs = Presentation(str(source_path))
    target_month_label = build_target_month_label(codigo_mes, language)
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    updated = replace_month_label(run.text, target_month_label)
                    if updated != run.text:
                        run.text = updated
    prs.save(str(target_path))


def find_best_existing_template(codigo_mes):
    base_code, suffix = split_codigo_mes(codigo_mes)
    suffix_aliases = [suffix] if suffix else []
    if suffix and suffix.upper() in ("MX", "MEX"):
        suffix_aliases = ["MX", "MEX"]

    exact_candidates = []
    if codigo_mes:
        exact_candidates.extend([
            paths.PPTX_DIR / f"{codigo_mes}.pptx",
            paths.PPTX_DIR / f"{codigo_mes.replace('-', ' - ')}.pptx",
        ])
    if suffix_aliases:
        for alias in suffix_aliases:
            exact_candidates.extend([
                paths.PPTX_DIR / f"{base_code}-{alias}.pptx",
                paths.PPTX_DIR / f"{base_code} - {alias}.pptx",
            ])

    for candidate in exact_candidates:
        if candidate.exists():
            return candidate

    template_pairs = []
    for path in paths.PPTX_DIR.glob("*.pptx"):
        stem = path.stem
        if suffix_aliases:
            if re.match(r"^\d{6}\s*-\s*(?:mx|mex)$", stem, re.IGNORECASE):
                template_pairs.append(path)
        else:
            if re.match(r"^\d{6}$", stem):
                template_pairs.append(path)

    if not template_pairs and suffix_aliases:
        template_pairs = [path for path in paths.PPTX_DIR.glob("*.pptx") if re.match(r"^\d{6}$", path.stem, re.IGNORECASE)]

    if not template_pairs and not suffix_aliases:
        template_pairs = [path for path in paths.PPTX_DIR.glob("*.pptx") if re.match(r"^\d{6}(?:\s*-\s*(?:mx|mex))?$", path.stem, re.IGNORECASE)]

    if not template_pairs:
        return None

    target_num = int(base_code)

    def template_score(path):
        match = re.match(r"^(\d{6})", path.stem)
        if not match:
            return (10**9, path.name)
        return (abs(int(match.group(1)) - target_num), path.name)

    return sorted(template_pairs, key=template_score)[0]


def parse_codigo_mes_arg(arg):
    arg = arg.strip()
    match = re.match(r"^(\d{4})\s*[-]?\s*(\d{2})(?:\s*[-\s]?\s*([A-Za-z]{2,3}))?$", arg, re.IGNORECASE)
    if match:
        ano, mes, suffix = match.groups()
        return build_codigo_mes(ano, mes, suffix)
    return None


def mapear_cliente(nome):
    nome_norm = normalizar(nome)
    if nome_norm in CLIENT_MAP:
        return CLIENT_MAP[nome_norm]
    nome_simplified = re.sub(r"[\s\-_]+", "", nome_norm)
    for key in sorted(CLIENT_MAP.keys(), key=len, reverse=True):
        key_norm = normalizar(key)
        if key_norm and key_norm in nome_norm:
            return CLIENT_MAP[key]
        if nome_norm and nome_norm in key_norm:
            return CLIENT_MAP[key]
        key_simplified = re.sub(r"[\s\-_]+", "", key_norm)
        if key_simplified and key_simplified in nome_simplified:
            return CLIENT_MAP[key]
        if nome_simplified and nome_simplified in key_simplified:
            return CLIENT_MAP[key]
    return nome_norm


def selecionar_codigo_mes(argv=None):
    if argv is None:
        argv = sys.argv[1:]

    if len(argv) >= 1:
        codigo = parse_codigo_mes_arg(argv[0])
        if codigo and len(argv) == 1:
            return codigo
        if len(argv) >= 2:
            ano = re.sub(r"\D", "", argv[0])
            segundo = argv[1].strip()
            third = argv[2].strip() if len(argv) >= 3 else ""
            suffix = segundo.strip().lstrip("-")
            suffix_terceiro = third.strip().lstrip("-")
            if suffix.upper() in ("MX", "MEX"):
                if len(ano) == 4 and len(argv) >= 3:
                    mes = re.sub(r"\D", "", argv[2])
                    if len(mes) in (1, 2):
                        return build_codigo_mes(ano, mes, suffix)
                codigo_6 = re.sub(r"\D", "", argv[0])
                if len(codigo_6) == 6:
                    return build_codigo_mes(codigo_6[:4], codigo_6[4:], suffix)
            if suffix_terceiro.upper() in ("MX", "MEX") and len(ano) == 4 and len(segundo) in (1, 2) and segundo.isdigit():
                return build_codigo_mes(ano, segundo, suffix_terceiro)
            if len(ano) == 4 and len(segundo) in (1, 2) and segundo.isdigit():
                return ano + segundo.zfill(2)

    while True:
        entrada = input("Digite ano e mês para gerar o PPT (ex: 2026 02, 202602 ou 202602 MX): ").strip()
        codigo = parse_codigo_mes_arg(entrada)
        if codigo:
            return codigo
        print("Formato inválido. Use 202602, 2026 02 ou 202602 MX.")


def listar_templates_ppt(base_dir):
    try:
        arquivos = [f for f in os.listdir(base_dir) if f.lower().endswith((".pptx", ".odp"))]
    except FileNotFoundError:
        return []
    codigos = []
    for f in arquivos:
        stem = os.path.splitext(f)[0]
        if re.match(r"^\d{6}(?:\s*[-]\s*(?:mx|mex))?$", stem, re.IGNORECASE):
            codigos.append(stem)
    return sorted(codigos)


def split_codigo_mes(codigo_mes):
    codigo_norm = parse_codigo_mes_arg(codigo_mes)
    if not codigo_norm:
        return codigo_mes, None
    if is_spanish_template_stem(codigo_norm):
        base_code, suffix = codigo_norm.split("-", 1)
        return base_code, suffix.upper()
    return codigo_norm, None


def build_mes_code(mes_norm):
    if not mes_norm:
        return None
    parts = mes_norm.split("-")
    if len(parts) < 2:
        return None
    month_name = parts[0].strip()
    year = parts[-1].strip()
    month_num = MONTH_NAME_TO_NUMBER.get(month_name)
    if month_num is None:
        return None
    return f"{year}-{month_num:02d}"


def find_soffice_executable():
    candidates = [
        shutil.which("soffice"),
        shutil.which("libreoffice"),
    ]
    if os.name == "nt":
        program_files = os.environ.get("PROGRAMFILES", r"C:\Program Files")
        program_files_x86 = os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)")
        candidates.extend([
            str(Path(program_files) / "LibreOffice" / "program" / "soffice.exe"),
            str(Path(program_files_x86) / "LibreOffice" / "program" / "soffice.exe"),
            str(Path(program_files) / "LibreOffice 7.0" / "program" / "soffice.exe"),
            str(Path(program_files_x86) / "LibreOffice 7.0" / "program" / "soffice.exe"),
        ])
    for candidate in candidates:
        if candidate and Path(candidate).exists():
            return str(candidate)
    return None


def convert_odp_to_pptx(odp_path, output_dir):
    soffice = find_soffice_executable()
    if not soffice:
        raise FileNotFoundError(
            "LibreOffice/soffice não encontrado. Instale LibreOffice para converter .odp para .pptx."
        )
    output_path = output_dir / f"{odp_path.stem}.pptx"
    args = [
        soffice,
        "--headless",
        "--convert-to",
        "pptx",
        "--outdir",
        str(output_dir),
        str(odp_path),
    ]
    subprocess.run(args, check=True)
    if not output_path.exists():
        raise FileNotFoundError(f"Falha ao converter {odp_path} para PPTX.")
    return output_path


def find_fallback_pptx(codigo_mes):
    base_code, suffix = split_codigo_mes(codigo_mes)
    candidates = []
    if codigo_mes:
        candidates.extend([
            paths.PPTX_DIR / f"{codigo_mes}.pptx",
            paths.PPTX_DIR / f"{codigo_mes.replace('-', ' - ')}.pptx",
        ])
    if suffix:
        candidates.extend([
            paths.PPTX_DIR / f"{base_code}.pptx",
            paths.PPTX_DIR / f"{base_code}-{suffix}.pptx",
            paths.PPTX_DIR / f"{base_code} - {suffix}.pptx",
        ])
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return None


def find_template_file(codigo_mes):
    base_code, suffix = split_codigo_mes(codigo_mes)
    language = "es" if suffix and suffix.upper() in ("MX", "MEX") else "pt"
    target_name = f"{base_code}.pptx" if not suffix else f"{base_code} - {suffix}.pptx"
    target_path = paths.PPTX_DIR / target_name
    if target_path.exists():
        return target_path

    source_path = find_best_existing_template(codigo_mes)
    if source_path is None:
        return None

    target_path.parent.mkdir(parents=True, exist_ok=True)
    duplicate_template_with_month(source_path, target_path, codigo_mes, language)
    print(f"Template criado a partir de {source_path.name}: {target_path.name}")
    return target_path


def executar_todos(base_dir):
    codigos = listar_templates_ppt(base_dir)
    if not codigos:
        print(f"Nenhum template .pptx ou .odp no diretório {base_dir} encontrado.")
        sys.exit(1)

    base_codes = sorted({split_codigo_mes(codigo)[0] for codigo in codigos})
    for codigo in base_codes:
        for variante in build_mes_variants(codigo):
            print(f"\n--- Gerando {variante} ---")
            gerar_ppt(variante)

    print("\nTodos os PPTs foram gerados.")
    sys.exit(0)


def build_mes_variants(codigo_mes):
    base_code, _ = split_codigo_mes(codigo_mes)
    return [base_code, f"{base_code}-MEX"]


def gerar_ppt(codigo_mes, target_client=None):
    print(f"Gerando PPT para: {codigo_mes}")

    excel_data = load_excel_reports()
    df_logs = load_throttling_logs()
    df_incidents = load_cloud_incidents()

    ppt_file = find_template_file(codigo_mes)
    if ppt_file is None:
        raise FileNotFoundError(f"Template não encontrado: {codigo_mes}.pptx")

    final_dir = paths.OUTPUT_PPTX_DIR / f"{codigo_mes}_final2"
    final_dir.mkdir(parents=True, exist_ok=True)
    final_ppt = final_dir / f"{codigo_mes}_final2.pptx"
    for caminho in [
        paths.OUTPUT_PPTX_DIR / f"{codigo_mes}_preenchido.pptx",
        paths.OUTPUT_PPTX_DIR / f"{codigo_mes}_final.pptx",
        final_ppt,
    ]:
        if caminho.exists():
            caminho.unlink()

    if not ppt_file.exists():
        raise FileNotFoundError(f"Template não encontrado: {ppt_file}")

    language = "es" if is_spanish_template_stem(codigo_mes) else "pt"
    if ppt_file.suffix.lower() == ".odp":
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                tmp_ppt = convert_odp_to_pptx(ppt_file, Path(tmpdir))
                prs = Presentation(str(tmp_ppt))
                fill_presentation(prs, excel_data, df_logs, df_incidents, codigo_mes, language)
                prs.save(str(final_ppt))
        except FileNotFoundError as e:
            print(f"Aviso: {e}")
            fallback = find_fallback_pptx(codigo_mes)
            if fallback:
                print(f"Usando fallback PPTX: {fallback}")
                prs = Presentation(str(fallback))
                fill_presentation(prs, excel_data, df_logs, df_incidents, codigo_mes, language)
                prs.save(str(final_ppt))
            else:
                raise
    else:
        prs = Presentation(str(ppt_file))
        fill_presentation(prs, excel_data, df_logs, df_incidents, codigo_mes, language)
        prs.save(str(final_ppt))

    clientes = split_presentation_by_client(final_ppt, final_dir, codigo_mes, target_client)

    print("\nPPT FINAL GERADO:")
    print(final_ppt)
    print("Diretório de saída:", final_dir)
    print("Clientes gerados:", ", ".join(clientes))

    return final_ppt, final_dir, clientes


def selecionar_modo(argv=None):
    if argv is None:
        argv = sys.argv[1:]

    if argv:
        primeiro = argv[0].strip().lower()
        if primeiro in ("todos", "all", "t"):
            return "all", None
        if primeiro in ("mx", "mex") and len(argv) >= 2:
            codigo = selecionar_codigo_mes(argv[1:])
            if codigo:
                if is_spanish_template_stem(codigo):
                    return "single", codigo
                return "single", build_codigo_mes(codigo[:4], codigo[4:6], primeiro.upper())
        codigo = selecionar_codigo_mes(argv)
        if codigo:
            return "single", codigo

    while True:
        escolha = input("Digite 1 para gerar um PPT ou 2 para gerar todos os PPTs disponíveis: ").strip().lower()
        if escolha in ("1", "single", "um", "um ppt", "m", "mes"):
            return "single", selecionar_codigo_mes()
        if escolha in ("2", "todos", "all", "t"):
            return "all", None
        print("Opção inválida. Digite 1 para um ou 2 para todos.")


def find_storage_csv_path(year):
    candidate = paths.INPUTS_DIR / f"Cloud - Storage {year}.csv"
    return candidate if candidate.exists() else None


def generate_excel_report(year):
    storage_path = find_storage_csv_path(year)
    if storage_path is None:
        raise FileNotFoundError(f"Arquivo de storage não encontrado para {year} em {paths.INPUTS_DIR}.")

    df_storage = pd.read_csv(storage_path, header=[0, 1], encoding="latin1")
    if "Customer" not in df_storage.columns.get_level_values(1):
        raise ValueError(f"{storage_path.name} sem coluna obrigatória: Customer")
    colunas = []
    mes_atual = ""
    for mes, tipo in df_storage.columns:
        if "Unnamed" not in mes:
            mes_atual = mes
        colunas.append((mes_atual, tipo))
    df_storage.columns = pd.MultiIndex.from_tuples(colunas)

    resultado = []
    meses = set([col[0] for col in df_storage.columns])
    for _, row in df_storage.iterrows():
        cliente = row.xs("Customer", level=1).values[0]
        for mes in meses:
            usuarios = None
            s3 = None
            database = None
            for m, tipo in df_storage.columns:
                if m != mes:
                    continue
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
                        "Database (Gb)": database,
                    })

    df_final = pd.DataFrame(resultado)
    if df_final.empty:
        raise ValueError(f"Nenhum dado consolidado encontrado em {storage_path.name}.")
    df_final["Mês"] = pd.to_datetime(df_final["Mês"])
    df_final["Disponibilidade"] = df_final["Mês"].dt.days_in_month * 24
    df_final = df_final.sort_values(["Cliente", "Mês"])
    df_final["Mês"] = df_final["Mês"].dt.date

    df_final["Cliente"] = df_final["Cliente"].apply(normalizar).replace(CLIENT_MAP)

    df_cloud = load_cloud_incidents()
    df_cloud = df_cloud[df_cloud["Year"] == int(year)].copy() if not df_cloud.empty else df_cloud
    if not df_cloud.empty:
        df_cloud["Minutos Indisponibilidade"] = df_cloud["Total Time (Minutes)"].where(
            df_cloud["Consider Unavailability?"].str.lower() == "yes", 0
        )
        df_inc_clean = (
            df_cloud.groupby(["Customer_norm", "Year", "Month_num"], as_index=False)
            .agg(
                Incidentes=("Customer_norm", "size"),
                **{"Minutos Indisponibilidade": ("Minutos Indisponibilidade", "sum")},
            )
            .rename(columns={"Customer_norm": "Cliente"})
        )
        df_inc_clean["Mês"] = pd.to_datetime(
            df_inc_clean["Year"].astype(str)
            + "-"
            + df_inc_clean["Month_num"].astype(str).str.zfill(2)
            + "-01"
        ).dt.date
        df_inc_clean["Horas"] = pd.DatetimeIndex(df_inc_clean["Mês"]).days_in_month * 24
        df_inc_clean["Disponibilidade (%)"] = (
            (df_inc_clean["Horas"] - df_inc_clean["Minutos Indisponibilidade"] / 60)
            / df_inc_clean["Horas"]
            * 100
        )
        df_inc_clean = df_inc_clean[
            [
                "Cliente",
                "Mês",
                "Horas",
                "Minutos Indisponibilidade",
                "Incidentes",
                "Disponibilidade (%)",
            ]
        ]
    else:
        df_inc_clean = pd.DataFrame(
            columns=[
                "Cliente",
                "Mês",
                "Horas",
                "Minutos Indisponibilidade",
                "Incidentes",
                "Disponibilidade (%)",
            ]
        )

    df_final = pd.merge(
        df_final,
        df_inc_clean,
        on=["Cliente", "Mês"],
        how="left"
    )
    return df_final


def should_regenerate_report(df, year):
    if df is None or df.empty:
        return True
    try:
        meses = sorted(pd.DatetimeIndex(df["Mês"]).month.unique())
    except Exception:
        return True
    if year == "2026" and 4 not in meses:
        return True
    return False


def load_excel_reports():
    relatorios = {}
    for ano in ("2025", "2026"):
        path = paths.OUTPUT_REPORTS_DIR / f"Relatorio{ano}_Com_Incidentes9.xlsx"
        df = None
        storage_path = find_storage_csv_path(ano)
        cloud_incident_paths = sorted(paths.INPUTS_DIR.glob(f"Cloud Incidents(Registro {ano}*.csv"))
        source_paths = [p for p in [storage_path, *cloud_incident_paths] if p is not None]
        if path.exists() and not is_report_outdated(path, source_paths):
            df = pd.read_excel(path)
            ensure_columns(df, ["Cliente"], path.name)
            # Procurar pela coluna de mês com diferentes nomes possíveis (devido a encoding)
            mes_col = None
            for col in df.columns:
                if "mes" in col.lower() or "mês" in col.lower():
                    mes_col = col
                    break
            if mes_col is None:
                raise ValueError(f"{path.name} sem coluna de mês reconhecível.")
            if mes_col:
                df[mes_col] = pd.to_datetime(df[mes_col]).dt.date
                # Renomear para "Mês" se tiver outro nome
                if mes_col != "Mês":
                    df = df.rename(columns={mes_col: "Mês"})
            df["Cliente"] = df["Cliente"].apply(normalizar).replace(CLIENT_MAP)
        if should_regenerate_report(df, ano):
            df = generate_excel_report(ano)
            # Salvar com engine xlsxwriter para evitar problemas de encoding
            df.to_excel(path, index=False, engine='openpyxl')
            print(f"Relatório {ano} regenerado a partir de inputs: {path}")
        relatorios[ano] = df
    return relatorios


def load_throttling_logs():
    rows = []
    if not paths.LOGS_DIR.exists():
        return pd.DataFrame(rows)

    for arquivo in sorted(paths.LOGS_DIR.glob("*-throttling-log-*.csv")):
        match = re.search(r"(.*?)-throttling-log-(\d{4})-(\d{2})", arquivo.name)
        if not match:
            continue
        cliente = normalizar(match.group(1))
        cliente = CLIENT_MAP.get(cliente, cliente)
        ano = match.group(2)
        mes_num = int(match.group(3))
        if mes_num < 1 or mes_num > 12:
            continue
        mes_texto = list(MONTH_NAME_TO_NUMBER.keys())[mes_num - 1] + f"-{ano}"
        try:
            df = pd.read_csv(arquivo, encoding="utf-8")
        except Exception:
            try:
                df = pd.read_csv(arquivo, encoding="latin1")
            except Exception:
                continue
        if "QTD" not in df.columns or "FUNCAO" not in df.columns:
            continue
        df["QTD"] = pd.to_numeric(df["QTD"], errors="coerce").fillna(0)
        total_qtd = int(df["QTD"].sum())
        funcoes = df["FUNCAO"].dropna().astype(str).tolist()
        top_funcoes = [f.strip().capitalize() for f, _ in Counter(funcoes).most_common(5)]
        rows.append({
            "cliente": cliente,
            "mes": normalizar(mes_texto),
            "mes_code": f"{ano}-{mes_num:02d}",
            "total_qtd": total_qtd,
            "top5": top_funcoes,
        })
    return pd.DataFrame(rows)


def load_cloud_incidents():
    arquivos = sorted(paths.INPUTS_DIR.glob("Cloud Incidents(Registro*.csv"))
    dfs = []
    for arquivo in arquivos:
        try:
            df = pd.read_csv(arquivo, encoding="utf-8")
        except Exception:
            df = pd.read_csv(arquivo, encoding="latin1")
        ensure_columns(df, ["Customer", "Year", "Month"], arquivo.name)
        if "Total Time (Minutes)" in df.columns:
            df["Total Time (Minutes)"] = pd.to_numeric(df["Total Time (Minutes)"], errors="coerce").fillna(0)

        if "Consider Unavailability?" not in df.columns:
            for col in df.columns:
                if normalizar(col) == normalizar("Consider Unavailability?"):
                    df = df.rename(columns={col: "Consider Unavailability?"})
                    break
            else:
                for col in df.columns:
                    values = df[col].astype(str).str.strip().str.lower()
                    if values.isin(["yes", "no"]).any() and col.startswith("Unnamed"):
                        df = df.rename(columns={col: "Consider Unavailability?"})
                        break

        if "Consider Unavailability?" in df.columns:
            df["Consider Unavailability?"] = df["Consider Unavailability?"].fillna("No").astype(str).str.strip()
        else:
            df["Consider Unavailability?"] = "No"

        df["Customer_norm"] = df["Customer"].apply(normalizar).apply(mapear_cliente)
        df["Year"] = pd.to_numeric(df["Year"], errors="coerce").astype("Int64")
        df["Month_norm"] = df["Month"].apply(normalizar)
        df["Month_num"] = df["Month_norm"].map(MONTH_NAME_TO_NUMBER)
        df.loc[df["Month_norm"].str.isdigit(), "Month_num"] = df.loc[df["Month_norm"].str.isdigit(), "Month_norm"].astype(int)
        df["Month_num"] = df["Month_num"].astype("Int64")
        df["Mes_norm"] = df["Month_norm"] + "-" + df["Year"].astype(str)
        df["Mes_code"] = df["Year"].astype(str) + "-" + df["Month_num"].astype(str).str.zfill(2)

        df = df[
            df["Customer_norm"].astype(bool)
            & df["Month_norm"].astype(bool)
            & df["Year"].notna()
            & df["Month_num"].notna()
        ]
        dfs.append(df)
    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()


def remover_bullet(paragraph):
    try:
        from lxml import etree
        pPr = paragraph._element.get_or_add_pPr()
        
        # Remover elementos de bullet existentes
        ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
        for tag in ("buChar", "buAutoNum", "buBlip", "buFont", "buNone"):
            el = pPr.find(f"a:{tag}", ns)
            if el is not None:
                pPr.remove(el)
        
        # Remover espaço de indent que está associado a bullets
        for tag in ("marL", "indent"):
            el = pPr.find(f"a:{tag}", ns)
            if el is not None:
                pPr.remove(el)
        
        # Adicionar buNone para desabilitar bullets explicitamente
        bu_none = etree.Element("{http://schemas.openxmlformats.org/drawingml/2006/main}buNone")
        pPr.insert(0, bu_none)
    except Exception:
        pass


def clear_text_frame(text_frame):
    try:
        for paragraph in list(text_frame.paragraphs)[::-1]:
            p = paragraph._element
            p.getparent().remove(p)
    except Exception:
        pass
    text_frame.text = ""


def adicionar_linha_tabela(tabela, template_row_idx=1):
    if len(tabela.rows) <= template_row_idx:
        raise IndexError("Template row index not available for table row duplication.")
    template_row = tabela.rows[template_row_idx]
    new_tr = copy.deepcopy(template_row._tr)
    tabela._tbl.append(new_tr)
    new_row = tabela.rows[len(tabela.rows) - 1]
    for cell in new_row.cells:
        cell.text = ""
    return new_row


def fill_excel_metrics(slide, df_excel, df_incidents, cliente_final, mes_data, mes_norm, language="pt"):
    dados = df_excel[(df_excel["Cliente"] == cliente_final) & (df_excel["Mês"] == mes_data)]
    if dados.empty:
        mensagem = "Sem dados Excel" if language == "pt" else "Sin datos de Excel"
        print(f"{mensagem} para {cliente_final} - {mes_data}")
        return
    row = dados.iloc[0]

    def safe_int(valor):
        valor = pd.to_numeric(valor, errors="coerce")
        if pd.isna(valor):
            return 0
        return int(valor)

    def safe_percentage(valor):
        if pd.isna(valor) or valor == "":
            return None
        if isinstance(valor, str):
            texto = valor.strip()
            if not texto:
                return None
            if texto.endswith("%"):
                return texto
            try:
                return f"{float(texto.replace(",", ".")):.2f}%"
            except ValueError:
                return texto
        try:
            return f"{float(valor):.2f}%"
        except Exception:
            return str(valor)

    def format_storage_value(valor):
        texto = str(valor).strip()
        if not texto or texto.lower() in ("nan", "none"):
            return ""
        if re.search(r"(?:gb|tb|pb|mb)", texto.lower()):
            return texto
        return f"{texto} GB"

    def center_shape_text(shape):
        if not shape.has_text_frame:
            return
        for paragraph in shape.text_frame.paragraphs:
            paragraph.alignment = PP_ALIGN.CENTER

    valor_usuarios = safe_int(row.get("Usuários", 0))
    valor_incidentes = safe_int(row.get("Incidentes", 0))
    valor_indisp = safe_int(row.get("Minutos Indisponibilidade", 0))
    valor_db = format_storage_value(row.get("Database (Gb)", ""))
    valor_s3 = format_storage_value(row.get("S3 (Gb)", ""))
    valor_disponibilidade = row.get("Disponibilidade", "")
    valor_disponibilidade_pct = safe_percentage(row.get("Disponibilidade (%)", ""))

    mes_code = build_mes_code(mes_norm)
    incident_rows = df_incidents[
        (df_incidents["Customer_norm"] == cliente_final)
        & (
            (df_incidents["Mes_norm"] == mes_norm)
            | ((mes_code is not None) & (df_incidents["Mes_code"] == mes_code))
        )
    ]
    if not incident_rows.empty:
        valor_incidentes = int(len(incident_rows))
        valor_indisp = int(
            pd.to_numeric(
                incident_rows[incident_rows["Consider Unavailability?"].str.lower() == "yes"]["Total Time (Minutes)"],
                errors="coerce"
            ).fillna(0).sum()
        )

    horas_disponiveis = pd.to_numeric(row.get("Horas"), errors="coerce")
    disponibilidade_num = pd.to_numeric(row.get("Disponibilidade"), errors="coerce")
    total_hours = None
    if hasattr(mes_data, "year"):
        total_hours = calendar.monthrange(mes_data.year, mes_data.month)[1] * 24

    if total_hours is not None and total_hours > 0:
        if valor_indisp > 0:
            disponibilidade_num = total_hours - valor_indisp / 60.0
        elif pd.isna(disponibilidade_num) and not pd.isna(horas_disponiveis) and horas_disponiveis > 0:
            disponibilidade_num = horas_disponiveis
        elif pd.isna(disponibilidade_num):
            disponibilidade_num = total_hours

        if not pd.isna(disponibilidade_num):
            valor_disponibilidade_pct = f"{float(disponibilidade_num) / total_hours * 100:.2f}%"
            valor_disponibilidade = str(int(disponibilidade_num))

    if valor_disponibilidade_pct is None:
        valor_disponibilidade_pct = "Sem dados" if language == "pt" else "Sin datos"

    shapes = list(slide.shapes)
    for idx, shape in enumerate(shapes):
        if not shape.has_text_frame:
            continue
        texto = normalizar(shape.text)
        if any(keyword in texto for keyword in ["% disponibilidade", "% disponibilidad"] ) and idx + 1 < len(shapes):
            target = shapes[idx + 1]
            target.text = valor_disponibilidade_pct
            center_shape_text(target)
        elif any(keyword in texto for keyword in ["nro de incidentes", "numero de incidentes", "número de incidentes"]) and idx + 1 < len(shapes):
            target = shapes[idx + 1]
            target.text = str(valor_incidentes)
            center_shape_text(target)
        elif texto in ["disponibilidade", "disponibilidad"] and idx + 1 < len(shapes):
            target = shapes[idx + 1]
            target.text = f"{valor_disponibilidade} hrs"
            center_shape_text(target)
        elif texto in ["indisponibilidade", "indisponibilidad"] and idx + 1 < len(shapes):
            target = shapes[idx + 1]
            target.text = f"{valor_indisp} min"
            center_shape_text(target)
        elif texto in ["usuarios", "usuários"] and idx + 1 < len(shapes):
            target = shapes[idx + 1]
            target.text = str(valor_usuarios)
            center_shape_text(target)
        elif any(keyword in texto for keyword in ["armazenamento", "almacenamiento"]):
            db_shape = next((s2 for s2 in shapes[idx + 1:] if s2.has_text_frame), None)
            if db_shape:
                db_text = valor_db if valor_db and re.search(r"(?:gb|tb|pb|mb)", valor_db.lower()) else valor_db
                db_shape.text = db_text
                for p in db_shape.text_frame.paragraphs:
                    remover_bullet(p)
                    p.alignment = PP_ALIGN.CENTER
            s3_shape = next((s2 for s2 in shapes[idx + 2:] if s2.has_text_frame), None)
            if s3_shape:
                s3_text = valor_s3 if valor_s3 and re.search(r"(?:gb|tb|pb|mb)", valor_s3.lower()) else valor_s3
                s3_shape.text = s3_text
                for p in s3_shape.text_frame.paragraphs:
                    remover_bullet(p)
                    p.alignment = PP_ALIGN.CENTER


def fill_top5_logs(slide, df_logs, cliente_final, mes_norm, language="pt"):
    mes_code = build_mes_code(mes_norm)
    filtro_logs = df_logs[
        (df_logs["cliente"] == cliente_final)
        & (
            (df_logs["mes"] == mes_norm)
            | ((mes_code is not None) & (df_logs["mes_code"] == mes_code))
        )
    ]
    total_qtd = 0
    top_funcoes = []
    if not filtro_logs.empty:
        row = filtro_logs.iloc[0]
        total_qtd = int(row["total_qtd"])
        top_funcoes = row["top5"]

    keywords_normalized = [normalizar(kw) for kw in ["ações pró-ativas", "ações pro-ativas", "acciones proactivas", "acciones pro-activas"]]
    for i, shape in enumerate(slide.shapes):
        if not shape.has_text_frame:
            continue
        texto = normalizar(shape.text)
        if any(keyword in texto for keyword in keywords_normalized):
            label_top = shape.top
            label_left = shape.left
            target_shape = None
            candidates = []
            max_dx = 2500000
            max_dy = 2000000
            for next_shape in slide.shapes:
                if next_shape is shape or not next_shape.has_text_frame:
                    continue
                candidate_text = next_shape.text.strip()
                if not candidate_text.isdigit():
                    continue
                dx = abs(next_shape.left - label_left)
                dy = abs(next_shape.top - label_top)
                if dx <= max_dx and dy <= max_dy:
                    candidates.append((dx + dy, next_shape))
            if candidates:
                target_shape = min(candidates, key=lambda item: item[0])[1]
            else:
                max_vertical_distance = max(shape.height * 4, 2000000)
                for j in range(i + 1, min(i + 10, len(slide.shapes))):
                    next_shape = slide.shapes[j]
                    if not next_shape.has_text_frame:
                        continue
                    candidate_text = next_shape.text.strip()
                    if not candidate_text.isdigit():
                        continue
                    delta_top = next_shape.top - label_top
                    if delta_top < 0 or delta_top > max_vertical_distance:
                        continue
                    target_shape = next_shape
                    break
                if target_shape is None:
                    for j in range(i + 1, min(i + 10, len(slide.shapes))):
                        next_shape = slide.shapes[j]
                        if next_shape.has_text_frame and next_shape.text.strip().isdigit():
                            target_shape = next_shape
                            break
            if target_shape:
                target_shape.text = str(total_qtd)
        if "top 5" in texto:
            tf = shape.text_frame
            clear_text_frame(tf)
            p_title = tf.paragraphs[0]
            if language == "es":
                p_title.text = "Top 5 - sesiones finalizadas"
            else:
                p_title.text = "Top 5 - sessões encerradas"
            p_title.font.size = Pt(12)
            p_title.font.bold = True
            remover_bullet(p_title)
            if top_funcoes:
                for f in top_funcoes:
                    p = tf.add_paragraph()
                    p.text = f
                    p.font.size = Pt(11)
                    p.alignment = PP_ALIGN.LEFT
            else:
                p = tf.add_paragraph()
                p.text = "Sin registros" if language == "es" else "Sem registros"
                p.font.size = Pt(11)
                p.alignment = PP_ALIGN.LEFT
                remover_bullet(p)


def fill_incident_table(slide, df_incidents, cliente_final, mes_norm, language="pt"):
    mes_code = build_mes_code(mes_norm)
    filtro = df_incidents[
        (df_incidents["Customer_norm"] == cliente_final)
        & (
            (df_incidents["Mes_norm"] == mes_norm)
            | ((mes_code is not None) & (df_incidents["Mes_code"] == mes_code))
        )
    ]
    filtro = filtro.copy()
    blank_fields = filtro[["Date", "Problem Description", "Correction Action", "Starting Date/Time", "Ending Date/Time"]].fillna("").astype(str)
    blank_fields = blank_fields.apply(lambda col: col.str.strip().str.lower().replace("nan", ""))
    empty_row = blank_fields.eq("").all(axis=1) & (filtro["Total Time (Minutes)"] == 0)
    filtro = filtro[~empty_row]

    filtro_yes = filtro[filtro["Consider Unavailability?"].str.lower() == "yes"]
    filtro_no = filtro[filtro["Consider Unavailability?"].str.lower() != "yes"]

    def format_value(value):
        if pd.isna(value):
            return ""
        return str(value)

    slide_contains_no_impact_section = any(
        shape.has_text_frame and any(keyword in normalizar(shape.text) for keyword in ["sem impacto", "sin impacto"])
        for shape in slide.shapes
    )

    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        titulo = normalizar(shape.text)
        if all(keyword not in titulo for keyword in ["detalhamento de incidente", "detalle de incidente", "detalle de incidentes"]):
            continue
        tabela = None
        for s2 in slide.shapes:
            if getattr(s2, "has_table", False) and s2.has_table and s2.top > shape.top:
                tabela = s2.table
                break
        if tabela is None:
            continue
        for cell in tabela.rows[0].cells:
            cell_text_norm = normalizar(cell.text)
            if "solu" in cell_text_norm:
                if language == "es":
                    cell.text = "Solución"
                for paragraph in cell.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(8)
                        run.font.color.rgb = RGBColor(255, 255, 255)
                        run.font.bold = True
            if language == "es" and cell_text_norm in ["fim"]:
                cell.text = "Fin"
                for paragraph in cell.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(8)
                        run.font.color.rgb = RGBColor(255, 255, 255)
                        run.font.bold = True
        if any(keyword in titulo for keyword in ["sem impacto", "sin impacto"]):
            dados = filtro_no
        elif slide_contains_no_impact_section:
            dados = filtro_yes
        else:
            dados = filtro
        for r in range(1, len(tabela.rows)):
            for c in tabela.rows[r].cells:
                c.text = ""
        if dados.empty:
            if len(tabela.rows) > 1:
                row = tabela.rows[1]
                row.cells[0].text = "Resumen" if language == "es" else "Resumo"
                row.cells[5].text = "Sin incidentes en el mes." if language == "es" else "Sem registros de incidentes no mês."
                # Aplicar formatação de fonte para as células de mensagem vazia
                for cell in [row.cells[0], row.cells[5]]:
                    for paragraph in cell.text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.font.size = Pt(8)
            continue
        while len(tabela.rows) < len(dados) + 1:
            adicionar_linha_tabela(tabela)
        for i, (_, inc) in enumerate(dados.iterrows(), start=1):
            row = tabela.rows[i]
            row.cells[0].text = format_value(inc.get("Date", ""))
            row.cells[1].text = "Tasy"
            row.cells[2].text = format_value(inc.get("Starting Date/Time", inc.get("Start", "")))
            row.cells[3].text = format_value(inc.get("Ending Date/Time", inc.get("End", "")))
            row.cells[4].text = str(int(inc.get("Total Time (Minutes)", 0))) + " min"
            problema_text = format_value(inc.get("Problem Description", inc.get("Problem", "")))
            solucion_text = format_value(inc.get("Correction Action", inc.get("Solução", "")))
            row.cells[5].text = problema_text
            row.cells[6].text = solucion_text
        for r in range(1, len(tabela.rows)):
            for c in tabela.rows[r].cells:
                for paragraph in c.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(6)


def fill_presentation(prs, excel_data, df_logs, df_incidents, codigo_mes, language="pt"):
    for slide in prs.slides:
        texto_slide = " ".join([shape.text for shape in slide.shapes if shape.has_text_frame])
        cliente, mes_texto = extrair_cliente_mes(texto_slide)
        if not cliente or not mes_texto:
            continue
        ano_template = codigo_mes[:4]
        mes_texto = re.sub(r"-\d{4}$", f"-{ano_template}", mes_texto)
        cliente_final = mapear_cliente(cliente)
        mes_nome, ano = mes_texto.split("-")
        mes_num = MONTH_NAME_TO_NUMBER.get(normalizar(mes_nome))
        if not mes_num:
            print(f"Mês inválido no slide: {mes_texto}")
            continue
        mes_data = pd.to_datetime(f"{ano}-{mes_num:02d}-01").date()
        mes_norm = normalizar(f"{mes_nome}-{ano}")

        df_excel = excel_data.get(ano)
        if df_excel is None:
            print(f"Relatório não carregado para o ano {ano}")
            continue

        fill_excel_metrics(slide, df_excel, df_incidents, cliente_final, mes_data, mes_norm, language)
        fill_top5_logs(slide, df_logs, cliente_final, mes_norm, language)
        fill_incident_table(slide, df_incidents, cliente_final, mes_norm, language)


def extract_slide_client(slide):
    texto_slide = " ".join([shape.text for shape in slide.shapes if shape.has_text_frame])
    cliente, _ = extrair_cliente_mes(texto_slide)
    return mapear_cliente(cliente) if cliente else None


def duplicate_slide(prs, slide):
    blank_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[0]
    new_slide = prs.slides.add_slide(blank_layout)
    for shape in list(new_slide.shapes):
        sp = shape.element
        sp.getparent().remove(sp)
    for shape in slide.shapes:
        new_slide.shapes._spTree.insert_element_before(copy.deepcopy(shape.element), 'p:extLst')
    return new_slide


def sanitize_filename(text):
    safe = re.sub(r"[^a-zA-Z0-9_-]", "_", text.strip().lower())
    return re.sub(r"_+", "_", safe).strip("_") or "cliente"


def remove_slide(prs, slide):
    slide_id = str(slide.slide_id)
    for sldId in list(prs.slides._sldIdLst):
        if sldId.get("id") == slide_id:
            prs.slides._sldIdLst.remove(sldId)
            break


def split_presentation_by_client(final_ppt_path, output_dir, codigo_mes, target_client=None):
    source_prs = Presentation(str(final_ppt_path))
    clientes = {}
    for idx, slide in enumerate(source_prs.slides):
        cliente = extract_slide_client(slide)
        if not cliente:
            continue
        clientes.setdefault(cliente, []).append(idx)

    for cliente, slide_indices in clientes.items():
        if target_client and cliente != target_client:
            continue
        client_prs = Presentation(str(final_ppt_path))
        slide_indices_set = set([0] + slide_indices)  # Incluir o primeiro slide (índice 0) em todos os PPTs individuais
        for idx, slide in reversed(list(enumerate(client_prs.slides))):
            if idx not in slide_indices_set:
                remove_slide(client_prs, slide)
        filename = output_dir / f"{codigo_mes}_{sanitize_filename(cliente)}.pptx"
        client_prs.save(str(filename))
    return list(clientes.keys())


def main():
    paths.ensure_directories()
    modo, codigo_mes = selecionar_modo()

    if modo == "all":
        executar_todos(paths.PPTX_DIR)

    codigo_mes = codigo_mes or selecionar_codigo_mes()
    for variante in build_mes_variants(codigo_mes):
        gerar_ppt(variante)


if __name__ == "__main__":
    main()

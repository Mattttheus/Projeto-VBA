"""Back-end de processamento em lote do Casa Organizada.

Chamado pelo VBA da planilha. Recebe os dados das tabelas por arquivos temporários,
processa com pandas e devolve o resultado para o VBA gravar nas mesmas tabelas.

Comandos:
  importar  lê um extrato (CSV, XLSX ou OFX), categoriza pelas regras, remove duplicados
  auditar   procura duplicidades, valores atípicos, campos vazios e variações por categoria
"""
from __future__ import annotations

import argparse
import re
import sys
import unicodedata
from pathlib import Path
import pandas as pd

SEP = ";"
ENC = "cp1252"


def normalizar(texto: object) -> str:
    s = unicodedata.normalize("NFKD", str(texto or "")).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s).strip().upper()


def ler_tabela(caminho: str) -> pd.DataFrame:
    return pd.read_csv(caminho, sep=SEP, encoding=ENC, dtype=str, keep_default_na=False)


def numero(serie: pd.Series) -> pd.Series:
    """Converte '1.234,56', '1234.56', '-R$ 10,00' em float."""
    s = serie.astype(str).str.replace(r"[R$\s]", "", regex=True)
    tem_virgula = s.str.contains(",", regex=False)
    s = s.where(~tem_virgula, s.str.replace(".", "", regex=False).str.replace(",", ".", regex=False))
    return pd.to_numeric(s, errors="coerce")


# ------------------------------------------------------------------ extratos
COL_DATA = ["data", "date", "data lancamento", "data da transacao", "dt", "data movimento"]
COL_DESC = ["descricao", "historico", "title", "estabelecimento", "lancamento", "memo", "detalhes", "descricao da transacao"]
COL_VALOR = ["valor", "amount", "quantia", "valor (r$)", "valor r$", "montante"]


def _achar(colunas: list[str], candidatos: list[str]) -> str | None:
    norm = {normalizar(c).lower(): c for c in colunas}
    for cand in candidatos:
        if cand in norm:
            return norm[cand]
    for cand in candidatos:
        for chave, original in norm.items():
            if chave.startswith(cand):
                return original
    return None


def ler_ofx(caminho: Path) -> pd.DataFrame:
    txt = caminho.read_text(encoding="latin-1", errors="ignore")
    linhas = []
    for bloco in re.findall(r"<STMTTRN>(.*?)</STMTTRN>", txt, flags=re.S | re.I):
        def tag(nome: str) -> str:
            m = re.search(rf"<{nome}>([^<\r\n]+)", bloco, flags=re.I)
            return m.group(1).strip() if m else ""
        linhas.append({"data": tag("DTPOSTED")[:8], "descricao": tag("MEMO") or tag("NAME"), "valor": tag("TRNAMT")})
    df = pd.DataFrame(linhas)
    if not df.empty:
        df["data"] = pd.to_datetime(df["data"], format="%Y%m%d", errors="coerce")
        df["valor"] = pd.to_numeric(df["valor"], errors="coerce")
    return df


def ler_extrato(caminho: Path) -> pd.DataFrame:
    ext = caminho.suffix.lower()
    if ext == ".ofx":
        df = ler_ofx(caminho)
    else:
        if ext in (".xlsx", ".xls"):
            bruto = pd.read_excel(caminho, dtype=str)
        else:
            amostra = caminho.read_bytes()[:4096].decode("utf-8", errors="ignore")
            sep = ";" if amostra.count(";") > amostra.count(",") else ","
            try:
                bruto = pd.read_csv(caminho, sep=sep, dtype=str, encoding="utf-8-sig", keep_default_na=False)
            except UnicodeDecodeError:
                bruto = pd.read_csv(caminho, sep=sep, dtype=str, encoding="latin-1", keep_default_na=False)
        cd, cs, cv = (_achar(list(bruto.columns), c) for c in (COL_DATA, COL_DESC, COL_VALOR))
        if not (cd and cs and cv):
            raise ValueError(f"Colunas não reconhecidas no extrato: {list(bruto.columns)}")
        df = pd.DataFrame({
            "data": pd.to_datetime(bruto[cd], dayfirst=True, errors="coerce", format="mixed"),
            "descricao": bruto[cs].astype(str).str.strip(),
            "valor": numero(bruto[cv]),
        })
    df = df.dropna(subset=["data", "valor"])
    # conta corrente: saídas são negativas; fatura de cartão: todos positivos são gastos
    if (df["valor"] < 0).any():
        df["gasto"] = df["valor"] < 0
    else:
        df["gasto"] = df["valor"] > 0
    df["valor"] = df["valor"].abs().round(2)
    return df


def categorizar(desc: pd.Series, regras: pd.DataFrame, padrao: int) -> tuple[pd.Series, pd.Series]:
    regras = regras.assign(chave=regras["palavra"].map(normalizar)).query("chave != ''")
    regras = regras.assign(tam=regras["chave"].str.len()).sort_values("tam", ascending=False)
    pares = list(zip(regras["chave"], regras["codigo"].astype(int)))
    alvo = desc.map(normalizar)

    def achar(t: str) -> int | None:
        # a palavra-chave que aparece primeiro no texto vence (ex.: "IFOOD *RESTAURANTE" -> iFood); empate: a mais longa
        melhor = None
        for chave, cod in pares:
            pos = t.find(chave)
            if pos >= 0 and (melhor is None or pos < melhor[0]):
                melhor = (pos, cod)
        return melhor[1] if melhor else None

    cods = alvo.map(achar)
    revisar = cods.isna()
    return cods.fillna(padrao).astype(int), revisar


def cmd_importar(a: argparse.Namespace) -> None:
    extrato = ler_extrato(Path(a.extrato))
    lidos = len(extrato)
    creditos = int((~extrato["gasto"]).sum())
    df = extrato[extrato["gasto"]].copy()
    fora = int((df["data"].dt.year != a.ano).sum())
    df = df[df["data"].dt.year == a.ano]

    regras = ler_tabela(a.regras)
    codigos_validos = set(pd.to_numeric(ler_tabela(a.cats)["codigo"], errors="coerce").dropna().astype(int))
    regras = regras[pd.to_numeric(regras["codigo"], errors="coerce").isin(codigos_validos)]
    df["codigo"], df["revisar"] = categorizar(df["descricao"], regras, a.padrao)

    # duplicados: contra os lançamentos existentes e dentro do próprio extrato
    lanc = ler_tabela(a.lanc)
    chave = lambda d, v, t: d.dt.strftime("%Y-%m-%d") + "|" + v.map("{:.2f}".format) + "|" + t.map(normalizar).str.replace(r"^(CONTA: |APORTE: |\[REVISAR\] )", "", regex=True)
    existentes = set()
    if not lanc.empty:
        existentes = set(chave(pd.to_datetime(lanc["data"], errors="coerce"), numero(lanc["valor"]).fillna(0), lanc["descricao"]))
    df["chave"] = chave(df["data"], df["valor"], df["descricao"])
    duplicados = int(df["chave"].isin(existentes).sum() + df.duplicated("chave").sum())
    df = df[~df["chave"].isin(existentes)].drop_duplicates("chave")

    df["descricao"] = df["descricao"].str.replace(SEP, ",", regex=False).str.slice(0, 120)
    df.loc[df["revisar"], "descricao"] = "[revisar] " + df.loc[df["revisar"], "descricao"]
    saida = df.sort_values("data")[["data", "codigo", "descricao", "valor"]]
    saida.assign(data=saida["data"].dt.strftime("%Y-%m-%d"), valor=saida["valor"].map("{:.2f}".format)) \
        .to_csv(a.saida, sep=SEP, index=False, header=False, encoding=ENC, errors="replace")

    print(f"lidos={lidos}")
    print(f"novos={len(saida)}")
    print(f"duplicados={duplicados}")
    print(f"revisar={int(df['revisar'].sum())}")
    print(f"creditos={creditos}")
    print(f"fora_do_ano={fora}")


# ------------------------------------------------------------------ auditoria
def cmd_auditar(a: argparse.Namespace) -> None:
    df = ler_tabela(a.lanc)
    if df.empty:
        print("total=0")
        return
    df["linha"] = range(1, len(df) + 1)
    df["valor_n"] = numero(df["valor"])
    df["data_d"] = pd.to_datetime(df["data"], errors="coerce")

    dup = df.duplicated(["data", "codigo", "valor"], keep=False) & df["data"].ne("")
    duplicados = df.loc[dup, "linha"].tolist()

    grp = df.groupby("subcategoria")["valor_n"]
    mediana = grp.transform("median")
    qtd = grp.transform("count")
    atip = (qtd >= 4) & (df["valor_n"] > 3 * mediana) & (df["valor_n"] - mediana > 100)
    atipicos = df.loc[atip & ~dup, "linha"].tolist()

    sem_pag = int((df["pagamento"].str.strip() == "").sum())
    sem_resp = int((df["responsavel"].str.strip() == "").sum())

    # variação do último mês com gastos contra a média dos meses anteriores, por subcategoria
    df["mes"] = df["data_d"].dt.month
    piv = df.pivot_table(index="subcategoria", columns="mes", values="valor_n", aggfunc="sum", fill_value=0)
    variacoes = []
    if piv.shape[1] >= 2:
        ultimo = piv.columns.max()
        base = piv.drop(columns=ultimo).replace(0, pd.NA).mean(axis=1)
        var = ((piv[ultimo] - base) / base).dropna()
        var = var[(var.abs() >= 0.2) & (piv[ultimo] > 0)].sort_values(key=lambda s: -s.abs()).head(5)
        variacoes = [f"{nome} {v:+.0%}" for nome, v in var.items()]

    print(f"total={len(df)}")
    print(f"duplicados={','.join(map(str, duplicados))}")
    print(f"atipicos={','.join(map(str, atipicos))}")
    print(f"sem_pagamento={sem_pag}")
    print(f"sem_responsavel={sem_resp}")
    print(f"variacoes={' | '.join(variacoes)}")


def main() -> None:
    p = argparse.ArgumentParser(description=__doc__)
    sub = p.add_subparsers(dest="cmd", required=True)
    i = sub.add_parser("importar")
    i.add_argument("--extrato", required=True)
    i.add_argument("--lanc", required=True)
    i.add_argument("--regras", required=True)
    i.add_argument("--cats", required=True)
    i.add_argument("--saida", required=True)
    i.add_argument("--ano", type=int, required=True)
    i.add_argument("--padrao", type=int, required=True)
    i.set_defaults(func=cmd_importar)
    au = sub.add_parser("auditar")
    au.add_argument("--lanc", required=True)
    au.set_defaults(func=cmd_auditar)
    args = p.parse_args()
    try:
        args.func(args)
    except SystemExit:
        raise
    except Exception as e:  # o VBA mostra a mensagem
        print(f"erro={type(e).__name__}: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()

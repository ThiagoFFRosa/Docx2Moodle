import re
from docx import Document
from lxml import etree

HIFEN = r"[–—-]"
LETRA_Q = r"[Qq]uest[aã]o"
REG_TITULO = re.compile(
    rf"^\s*(\d+)\s*ª?\s*{LETRA_Q}\s*{HIFEN}\s*(?:T[oó]pico|Unidade)\s*(\d+)\s*$",
    re.I,
)

ROMANOS = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X"]

REG_LINHA_FINAL_COMANDO = re.compile(
    r"^\s*(?:"
    r"considerando .* assinale .*|"
    r"com base .* assinale .*|"
    r"assinale a alternativa correta:?|"
    r"assinale a opção correta:?|"
    r"assinale a alternativa que .*:?|"
    r"é correto o que se afirma.*|"
    r"indica corretamente .*|"
    r"podemos afirmar que estão corretos:?|"
    r"podemos afirmar que está correto:?|"
    r"estão corretos:?|"
    r"está correto:?"
    r")\s*$",
    re.I,
)

REG_GATILHO_ASSERTIVAS = re.compile(
    r"(?:"
    r".*\bconsidere as afirmativas(?: apresentadas| a seguir)?(?: .*?)?|"
    r".*\bconsidere as afirmações(?: apresentadas| a seguir)?(?: .*?)?|"
    r".*\banalise as afirmativas(?: apresentadas| a seguir)?(?: .*?)?|"
    r".*\banalise as afirmações(?: apresentadas| a seguir)?(?: .*?)?|"
    r".*\banalise cada um dos seguintes itens(?: .*?)?|"
    r".*\banalise os itens a seguir(?: .*?)?|"
    r".*\bjulgue as afirmativas(?: apresentadas| a seguir)?(?: .*?)?|"
    r".*\bavalie as afirmativas(?: apresentadas| a seguir)?(?: .*?)?"
    r")\.?\s*$",
    re.I,
)


def norm_ws(s: str) -> str:
    return (s or "").replace("\xa0", " ").replace("\u2003", " ").strip()


def clean_text(raw: str) -> str:
    t = raw or ""
    t = t.replace("\\p", "\n")
    t = t.replace("\r\n", "\n").replace("\r", "\n")
    t = re.sub(r"\n{3,}", "\n\n", t)
    lines = [line.strip() for line in t.split("\n")]
    lines = [line for line in lines if line != ""]
    t = "\n".join(lines)
    t = re.sub(r"[ \t]{2,}", " ", t)
    return norm_ws(t)


def _already_has_roman_prefix(line: str) -> bool:
    t = norm_ws(line)
    return bool(re.match(r"^(?:I|II|III|IV|V|VI|VII|VIII|IX|X)\.\s+", t, re.I))


def _looks_like_assertiva_line(line: str) -> bool:
    t = norm_ws(line)
    if not t:
        return False
    if REG_LINHA_FINAL_COMANDO.match(t):
        return False
    if _already_has_roman_prefix(t):
        return False
    if len(t) < 12:
        return False
    return True


def _contains_roman_token(text: str, roman: str) -> bool:
    t = norm_ws(text).upper()
    return bool(re.search(rf"(?<![A-ZÀ-ÖØ-Ý]){re.escape(roman)}(?![A-ZÀ-ÖØ-Ý])", t))


def _alternativas_indicam_assertivas_romanas(alternativas: dict) -> bool:
    """
    Reconstrói romanos somente se alguma alternativa tiver II ou III.
    Não usa I sozinho para evitar falso positivo.
    """
    if not alternativas:
        return False

    for texto in alternativas.values():
        if _contains_roman_token(texto, "II") or _contains_roman_token(texto, "III"):
            return True

    return False


def restore_roman_assertions(raw: str, alternativas: dict | None = None) -> str:
    text = clean_text(raw)
    if not text:
        return ""

    lines = [norm_ws(x) for x in text.split("\n") if norm_ws(x)]
    if not lines:
        return text

    if not _alternativas_indicam_assertivas_romanas(alternativas or {}):
        return "\n".join(lines)

    if any(_already_has_roman_prefix(line) for line in lines):
        return "\n".join(lines)

    trigger_idx = None
    for i, line in enumerate(lines):
        ll = norm_ws(line).lower()
        if (
            REG_GATILHO_ASSERTIVAS.search(line)
            or ("analise" in ll and "itens" in ll)
            or ("considere" in ll and "afirmativas" in ll)
            or ("considere" in ll and "afirmações" in ll)
        ):
            trigger_idx = i
            break

    if trigger_idx is None:
        return "\n".join(lines)

    before = lines[: trigger_idx + 1]
    after = lines[trigger_idx + 1 :]

    assertivas = []
    tail = []

    for line in after:
        if REG_LINHA_FINAL_COMANDO.match(line):
            tail.append(line)
            continue

        if _looks_like_assertiva_line(line) and not tail:
            assertivas.append(line)
        else:
            tail.append(line)

    if len(assertivas) < 2:
        return "\n".join(lines)

    rebuilt = before[:]
    for idx, item in enumerate(assertivas):
        romano = ROMANOS[idx] if idx < len(ROMANOS) else f"{idx + 1}"
        rebuilt.append(f"{romano}. {item}")

    rebuilt.extend(tail)
    return "\n".join(rebuilt)


def _sanitize_moodle_text(raw: str, alternativas: dict | None = None) -> str:
    t = clean_text(raw)
    t = restore_roman_assertions(t, alternativas=alternativas)
    t = re.sub(r"\(\s*\)", "(  )", t)
    return t


def text_to_html_paragraphs(raw: str) -> str:
    text = clean_text(raw)
    if not text:
        return "<p></p>"

    lines = [norm_ws(x) for x in text.split("\n") if norm_ws(x)]
    if not lines:
        return "<p></p>"

    parts = []
    current = []

    def flush():
        nonlocal current
        if current:
            parts.append("<p>" + "<br />".join(current) + "</p>")
            current = []

    in_assertivas = False

    for line in lines:
        is_assertiva = _already_has_roman_prefix(line)
        is_final = bool(REG_LINHA_FINAL_COMANDO.match(line))

        if is_assertiva and not in_assertivas:
            flush()
            in_assertivas = True

        if in_assertivas and not (is_assertiva or is_final):
            flush()
            in_assertivas = False

        current.append(line)

        if in_assertivas and is_final:
            flush()
            in_assertivas = False

    flush()
    return "".join(parts)


def is_letra_alt(s: str) -> bool:
    return norm_ws(s).upper() in {"A", "B", "C", "D", "E"}


def parse_title_to_nums(titulo: str):
    t = norm_ws(titulo)
    m = REG_TITULO.match(t)
    if not m:
        return None, None
    return int(m.group(1)), int(m.group(2))


def flatten_table(table) -> list[str]:
    out = []
    seen = set()
    for row in table.rows:
        row_vals = []
        for cell in row.cells:
            txt = cell.text or ""
            key = norm_ws(txt)
            if key in seen:
                continue
            seen.add(key)
            row_vals.append(txt)
        out.extend(row_vals)
    return out


def extract_difficulty(text: str) -> str | None:
    t = norm_ws(text)
    tl = t.lower()
    if re.search(r"\(\s*x\s*\)\s*f[áa]cil", tl):
        return "facil"
    if re.search(r"\(\s*x\s*\)\s*m[ée]dio", tl):
        return "medio"
    if re.search(r"\(\s*x\s*\)\s*dif[íi]cil", tl):
        return "dificil"
    return None


def cell_is_mark(x: str) -> bool:
    t = norm_ws(x)
    return t.lower() == "x" or t in {"✓", "✔", "☒", "☑"}


def row_has_label(norm_vals, *labels):
    vals = [v.lower() for v in norm_vals if v]
    targets = {label.lower() for label in labels}
    return any(v in targets for v in vals)


def row_join_nonempty(vals):
    return "\n".join(v for v in vals if norm_ws(v)).strip()


TABULAR_HEADER_PATTERNS = [
    r"item",
    r"valor(?:\s*\([^)]*\))?",
    r"tempo(?:\s*\([^)]*\))?",
    r"n\.?\s*[º°o]?\s*de\s*pacientes",
    r"frequ[êe]ncia(?:\s+absoluta|\s+acumulada)?",
    r"aluno",
    r"nota",
    r"ve[íi]culo",
    r"peso",
    r"consumo(?:\s+de\s+energia)?",
    r"xi",
    r"fi",
    r"total",
    r"fonte\s*:",
]


def _is_mostly_numeric_line(line: str) -> bool:
    nums = re.findall(r"\d+(?:[.,]\d+)?", line)
    if not nums:
        return False
    stripped = re.sub(r"\d+(?:[.,]\d+)?", "", line)
    stripped = re.sub(r"[%R$()\-–—/:;,.\s]", "", stripped)
    return len(stripped) <= 2


def _contains_tabular_pattern(text: str) -> bool:
    """
    Detecta indícios de questão dependente de tabela/listagem tabular.
    Heurística por sinais combinados para capturar blocos em linhas quebradas.
    """
    t = clean_text(text)
    if not t:
        return False

    lines = [norm_ws(x) for x in t.split("\n") if norm_ws(x)]
    if len(lines) < 4:
        return False

    short_lines = sum(1 for line in lines if len(line) <= 22)
    numeric_lines = sum(1 for line in lines if _is_mostly_numeric_line(line))
    header_lines = 0
    label_value_switches = 0
    has_fonte = False

    for idx, line in enumerate(lines):
        lower = line.lower()
        if any(re.search(rf"\b{pattern}\b", lower, re.I) for pattern in TABULAR_HEADER_PATTERNS):
            header_lines += 1
        if re.search(r"\bfonte\s*:", lower):
            has_fonte = True

        if idx > 0:
            prev_numeric = _is_mostly_numeric_line(lines[idx - 1])
            cur_numeric = _is_mostly_numeric_line(line)
            if prev_numeric != cur_numeric:
                label_value_switches += 1

    score = 0
    if short_lines >= 5:
        score += 1
    if numeric_lines >= 3:
        score += 1
    if header_lines >= 2:
        score += 2
    if label_value_switches >= 4:
        score += 1
    if has_fonte:
        score += 1

    return score >= 3


def _question_discard_reason(enunciado: str) -> str | None:
    """
    Marca como lixo questões que dependem de elementos não textuais
    ou estrutura tabular que não deve ser convertida automaticamente.
    """
    t = clean_text(enunciado).lower()
    if not t:
        return "enunciado_vazio"

    gatilhos_lixo = [
        "tabela abaixo",
        "gráfico abaixo",
        "imagem abaixo",
        "figura abaixo",
        "mapa abaixo",
        "quadro abaixo",
        "na figura abaixo",
        "no gráfico abaixo",
        "na tabela abaixo",
        "conforme o gráfico",
        "conforme a tabela",
        "a tabela se refere",
        "o resultado está apresentado no gráfico",
        "observe o gráfico",
        "observe a figura",
        "observe a tabela",
        "analise o gráfico",
        "analise a figura",
        "analise a tabela",
        "abaixo temos os dados",
        "os dados obtidos foram tabelados abaixo",
        "urna abaixo",
        "dados foram tabelados",
    ]

    for gatilho in gatilhos_lixo:
        if gatilho in t:
            return f"gatilho_textual:{gatilho}"

    if any(p in t for p in ["gráfico", "imagem", "figura", "mapa", "quadro"]):
        return "referencia_visual_generica"

    if "tabela" in t:
        return "referencia_tabela"

    if any(re.search(rf"\b{pattern}\b", t, re.I) for pattern in TABULAR_HEADER_PATTERNS):
        if _contains_tabular_pattern(enunciado):
            return "bloco_tabular_por_cabecalhos"

    if _contains_tabular_pattern(enunciado):
        return "bloco_tabular_heuristica"

    return None


def _question_should_be_discarded(enunciado: str) -> bool:
    return _question_discard_reason(enunciado) is not None


def extract_blocks_from_rows(table):
    rows = []
    for row in table.rows:
        vals = []
        seen = set()
        for cell in row.cells:
            txt = cell.text or ""
            key = norm_ws(txt)
            if key in seen:
                continue
            seen.add(key)
            vals.append(txt)
        rows.append(vals)

    title = ""
    difficulty = None
    enunciado = ""
    alternativas = {}
    gabarito = None
    justificativa = ""

    for vals in rows:
        norm_vals = [norm_ws(v) for v in vals]
        joined = " ".join(v for v in norm_vals if v)

        if not title:
            for v in norm_vals:
                if REG_TITULO.match(v):
                    title = v
                    break

        if difficulty is None and "nível de dificuldade" in joined.lower():
            difficulty = extract_difficulty(joined)

    idx_enunciado = None
    idx_alternativas = None
    idx_justificativa = None

    for idx, vals in enumerate(rows):
        norm_vals = [norm_ws(v) for v in vals]
        joined_lower = " ".join(v for v in norm_vals if v).lower()

        if idx_enunciado is None and row_has_label(norm_vals, "Enunciado"):
            idx_enunciado = idx

        if idx_alternativas is None and (
            row_has_label(norm_vals, "Alternativas")
            or "alternativas:" in joined_lower
            or joined_lower.strip() == "alternativas"
        ):
            idx_alternativas = idx

        if idx_justificativa is None and any(
            v.lower().startswith("justificativa") for v in norm_vals if v
        ):
            idx_justificativa = idx

    if idx_enunciado is not None:
        end_idx = idx_alternativas if idx_alternativas is not None else len(rows)
        enunciado_parts = []

        for j in range(idx_enunciado + 1, end_idx):
            block = row_join_nonempty(rows[j])
            if not norm_ws(block):
                continue
            enunciado_parts.append(block)

        enunciado = "\n".join(enunciado_parts).strip()

    if idx_alternativas is not None:
        end_idx = idx_justificativa if idx_justificativa is not None else len(rows)

        for j in range(idx_alternativas + 1, end_idx):
            raw_vals = rows[j]
            vals2 = [norm_ws(v) for v in raw_vals]
            vals2_nonempty = [v for v in vals2 if v]

            if not vals2_nonempty:
                continue

            joined2 = " ".join(vals2_nonempty).lower()
            if "justificativa" in joined2:
                break

            if len(vals2_nonempty) >= 2 and is_letra_alt(vals2_nonempty[0]):
                letra = vals2_nonempty[0].upper()

                texto_parts = []
                marcou = False

                for extra in vals2_nonempty[1:]:
                    if cell_is_mark(extra):
                        marcou = True
                    else:
                        texto_parts.append(extra)

                texto = "\n".join(texto_parts).strip()
                alternativas[letra] = texto

                if marcou:
                    gabarito = letra
                continue

            first = vals2_nonempty[0]
            if is_letra_alt(first):
                letra = first.upper()
                resto = "\n".join(vals2_nonempty[1:]).strip()
                if resto:
                    alternativas[letra] = resto

    if idx_justificativa is not None:
        justificativa_parts = []

        first_nonempty = [v for v in rows[idx_justificativa] if norm_ws(v)]
        if first_nonempty:
            row_text = "\n".join(first_nonempty)
            stripped = re.sub(r"(?is)^justificativa\s*:?\s*", "", row_text).strip()
            if stripped and stripped.lower() != row_text.strip().lower():
                justificativa_parts.append(stripped)

        for j in range(idx_justificativa + 1, len(rows)):
            block = row_join_nonempty(rows[j])
            if not norm_ws(block):
                continue
            justificativa_parts.append(block)

        justificativa = "\n".join(justificativa_parts).strip()

    return title, difficulty, enunciado, alternativas, gabarito, justificativa


def fallback_extract(flat_cells):
    base = [norm_ws(x) for x in flat_cells]
    enunciado = ""
    alternativas = {}
    gabarito = None
    justificativa = ""

    idx_enunciado = None
    idx_alternativas = None
    idx_justificativa = None

    for i, val in enumerate(base):
        vl = val.lower()

        if idx_enunciado is None and vl == "enunciado":
            idx_enunciado = i

        if idx_alternativas is None and (
            vl == "alternativas"
            or vl == "alternativas:"
            or "alternativas:" in vl
        ):
            idx_alternativas = i

        if idx_justificativa is None and vl.startswith("justificativa"):
            idx_justificativa = i

    if idx_enunciado is not None:
        end = idx_alternativas if idx_alternativas is not None else len(flat_cells)
        parts = []
        for i in range(idx_enunciado + 1, end):
            txt = flat_cells[i]
            if norm_ws(txt):
                parts.append(txt)
        enunciado = "\n".join(parts).strip()

    if idx_alternativas is not None:
        end = idx_justificativa if idx_justificativa is not None else len(base)
        i = idx_alternativas + 1
        while i < end:
            cur = base[i]

            if is_letra_alt(cur):
                letra = cur.upper()
                texto_parts = []
                marcou = False
                j = i + 1

                while j < end:
                    nxt = base[j]
                    if is_letra_alt(nxt):
                        break
                    if cell_is_mark(nxt):
                        marcou = True
                    elif nxt:
                        texto_parts.append(flat_cells[j])
                    j += 1

                alternativas[letra] = "\n".join(texto_parts).strip()
                if marcou:
                    gabarito = letra

                i = j
                continue

            i += 1

    if not gabarito:
        for i, val in enumerate(base):
            if cell_is_mark(val):
                for j in range(i - 1, -1, -1):
                    if is_letra_alt(base[j]):
                        gabarito = base[j].upper()
                        break
                if gabarito:
                    break

    if idx_justificativa is not None:
        parts = []

        first = flat_cells[idx_justificativa]
        stripped = re.sub(r"(?i)^justificativa\s*:?\s*", "", first).strip()
        if stripped:
            parts.append(stripped)

        for i in range(idx_justificativa + 1, len(flat_cells)):
            txt = flat_cells[i]
            if norm_ws(txt):
                parts.append(txt)

        justificativa = "\n".join(parts).strip()

    return enunciado, alternativas, gabarito, justificativa


def parse_docx_questions(docx_path: str):
    doc = Document(docx_path)
    questoes = []
    seq = 0

    for table in doc.tables:
        title, difficulty, enunciado, alternativas, gabarito, justificativa = extract_blocks_from_rows(table)

        if not title:
            flat = flatten_table(table)
            flat_norm = [norm_ws(t) for t in flat]
            titles = [t for t in flat_norm if REG_TITULO.match(t)]
            if not titles:
                continue
            title = titles[0]
            en_f, alt_f, gab_f, jus_f = fallback_extract(flat)
            enunciado = enunciado or en_f
            alternativas = alternativas or alt_f
            gabarito = gabarito or gab_f
            justificativa = justificativa or jus_f

        q_num, topico_num = parse_title_to_nums(title)
        if q_num is None:
            continue

        seq += 1

        alternativas_sanitizadas = {
            k: _sanitize_moodle_text(v) for k, v in alternativas.items()
        }

        enunciado_sanitizado = _sanitize_moodle_text(
            enunciado,
            alternativas=alternativas_sanitizadas,
        )

        discard_reason = _question_discard_reason(enunciado_sanitizado)
        discarded = discard_reason is not None

        questoes.append({
            "seq": seq,
            "questao_num": q_num,
            "topico_num": topico_num,
            "unidade_num": topico_num,
            "titulo": title,
            "dificuldade": difficulty,
            "enunciado": enunciado_sanitizado,
            "alternativas": alternativas_sanitizadas,
            "gabarito": gabarito,
            "justificativa": _sanitize_moodle_text(justificativa),
            "discarded": discarded,
            "discard_reason": discard_reason,
        })

    return questoes


def add_html_node(parent, tag, html):
    node = etree.SubElement(parent, tag, format="html")
    etree.SubElement(node, "text").text = etree.CDATA(html)
    return node


def make_moodle_xml(
    questoes,
    mark_remove_set=None,
    categoria="$module$/top/Padrão para AE1 - Atividade de Estudo 1",
    category_info="A categoria padrão para as questões compartilhadas no contexto 'AE1 - Atividade de Estudo 1'.",
    question_name_prefix="AE",
    skip_discarded=True,
):
    quiz = etree.Element("quiz")
    mark_remove_set = mark_remove_set or set()

    qcat = etree.SubElement(quiz, "question", type="category")
    category = etree.SubElement(qcat, "category")
    etree.SubElement(category, "text").text = categoria
    info = etree.SubElement(qcat, "info", format="moodle_auto_format")
    etree.SubElement(info, "text").text = category_info
    etree.SubElement(qcat, "idnumber").text = ""

    for q in questoes:
        if skip_discarded and q.get("discarded"):
            continue

        qnode = etree.SubElement(quiz, "question", type="multichoice")

        name = etree.SubElement(qnode, "name")
        display_name = f"{question_name_prefix} {q['seq']:02d}"

        if q.get("discarded") and not skip_discarded:
            display_name = "{LIXO} " + display_name
        elif (q["topico_num"], q["questao_num"]) in mark_remove_set:
            display_name = "{REMOVER} " + display_name

        etree.SubElement(name, "text").text = display_name

        add_html_node(qnode, "questiontext", text_to_html_paragraphs(q["enunciado"]))
        add_html_node(qnode, "generalfeedback", text_to_html_paragraphs(q.get("justificativa", "")))

        etree.SubElement(qnode, "defaultgrade").text = "1.0000000"
        etree.SubElement(qnode, "penalty").text = "0.3333333"
        etree.SubElement(qnode, "hidden").text = "0"
        etree.SubElement(qnode, "idnumber").text = ""
        etree.SubElement(qnode, "single").text = "true"
        etree.SubElement(qnode, "shuffleanswers").text = "true"
        etree.SubElement(qnode, "answernumbering").text = "abc"
        etree.SubElement(qnode, "showstandardinstruction").text = "0"

        add_html_node(qnode, "correctfeedback", "<p>Sua resposta está correta.</p>")
        add_html_node(qnode, "partiallycorrectfeedback", "<p>Sua resposta está parcialmente correta.</p>")
        add_html_node(qnode, "incorrectfeedback", "<p>Sua resposta está incorreta.</p>")
        etree.SubElement(qnode, "shownumcorrect")

        gab = (q.get("gabarito") or "").upper()
        for letra in ["A", "B", "C", "D", "E"]:
            alt_txt = text_to_html_paragraphs(q["alternativas"].get(letra, ""))
            fraction = "100" if gab == letra else "0"
            ans = etree.SubElement(qnode, "answer", fraction=fraction, format="html")
            etree.SubElement(ans, "text").text = etree.CDATA(alt_txt)
            fb = etree.SubElement(ans, "feedback", format="html")
            etree.SubElement(fb, "text").text = ""

    return etree.ElementTree(quiz)

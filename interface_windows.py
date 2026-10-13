"""Continua o DARF pela interface do Chrome comum, após a entrada manual."""

from __future__ import annotations

from decimal import Decimal, InvalidOperation
from pathlib import Path
from collections.abc import Callable
import ctypes
import re
import shutil
import time
import unicodedata
from urllib.parse import urlsplit


def _bibliotecas():
    try:
        import psutil
        import win32clipboard
        from pywinauto import Desktop
        from pywinauto.keyboard import send_keys
    except ImportError as exc:
        raise RuntimeError(
            "Instale as dependências do modo manual: python -m pip install -r requirements.txt"
        ) from exc
    return psutil, win32clipboard, Desktop, send_keys


def _texto(valor: str) -> str:
    sem_acentos = unicodedata.normalize("NFKD", valor)
    return "".join(c for c in sem_acentos if not unicodedata.combining(c)).casefold()


def _nome(controle) -> str:
    return controle.window_text().strip()


def _id(controle) -> str:
    try:
        return controle.automation_id().strip()
    except Exception:
        return ""


def _visivel(controle) -> bool:
    try:
        return controle.is_visible()
    except Exception:
        return False


def _elementos(janela, tipos: tuple[str, ...]):
    return [
        controle for controle in janela.descendants()
        if controle.element_info.control_type in tipos and _visivel(controle)
    ]


def _resumo_controles(janela) -> str:
    itens = []
    for controle in _elementos(janela, ("Edit", "ComboBox", "Button", "CheckBox", "ListItem")):
        nome = _nome(controle)
        nome = re.sub(r"\b\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\b", "[CNPJ]", nome)
        nome = re.sub(r"\b\d{14}\b", "[CNPJ]", nome)
        itens.append(f"{controle.element_info.control_type}: {nome[:70]!r} (id={_id(controle)!r})")
    return "; ".join(itens[:50]) or "nenhum controle acessível"


def _achar(janela, *, tipos: tuple[str, ...], ids: tuple[str, ...] = (),
           nomes: tuple[str, ...] = (), descricao: str):
    candidatos = _elementos(janela, tipos)
    por_id = [c for c in candidatos if _id(c) in ids]
    if len(por_id) == 1:
        return por_id[0]
    por_nome = [c for c in candidatos if any(n in _texto(_nome(c)) for n in nomes)]
    if len(por_nome) == 1:
        return por_nome[0]
    raise RuntimeError(
        f"Não foi possível identificar unicamente {descricao}. "
        f"Controles visíveis: {_resumo_controles(janela)}"
    )


def _valor(controle) -> str:
    try:
        return controle.get_value().strip()
    except Exception as exc:
        raise RuntimeError(f"Não foi possível ler o campo {_nome(controle)!r} pela interface do Windows.") from exc


def _digitar(controle, texto: str, send_keys, *, sair: bool = True) -> None:
    controle.click_input()
    send_keys("^a")
    send_keys(texto, with_spaces=True, pause=0.1)
    if sair:
        send_keys("{TAB}")


def _janela_ativa(psutil, Desktop):
    hwnd = ctypes.windll.user32.GetForegroundWindow()
    if not hwnd:
        raise RuntimeError("Não há uma janela ativa. Deixe a aba do Sicalc em primeiro plano.")
    janela = Desktop(backend="uia").window(handle=hwnd).wrapper_object()
    try:
        processo = psutil.Process(janela.process_id()).name().casefold()
    except Exception as exc:
        raise RuntimeError("Não foi possível identificar o processo da janela ativa.") from exc
    if processo != "chrome.exe" or janela.class_name() != "Chrome_WidgetWin_1":
        raise RuntimeError("A janela ativa não é o Google Chrome. Clique na aba do Sicalc e tente novamente.")
    return janela


def ativar_janela_sicalc() -> None:
    """Encontra uma guia do Sicalc no Chrome comum e a traz ao primeiro plano."""
    psutil, _, Desktop, _ = _bibliotecas()
    candidatos = []
    for janela in Desktop(backend="uia").windows():
        try:
            if (janela.class_name() != "Chrome_WidgetWin_1"
                    or psutil.Process(janela.process_id()).name().casefold() != "chrome.exe"):
                continue
            abas = [aba for aba in janela.descendants(control_type="TabItem")
                    if "sicalc" in _texto(_nome(aba))]
            if "sicalc" in _texto(_nome(janela)) or abas:
                candidatos.append((janela, abas))
        except Exception:
            continue
    if len(candidatos) != 1:
        raise RuntimeError(
            "Não encontrei uma única janela do Chrome com o Sicalc. "
            "Deixe somente a janela desejada aberta na tela de cálculo e tente novamente."
        )
    janela, abas = candidatos[0]
    janela.set_focus()
    if "sicalc" not in _texto(_nome(janela)):
        if len(abas) != 1:
            raise RuntimeError("Há mais de uma aba do Sicalc nesta janela; selecione a aba correta.")
        abas[0].select()
    time.sleep(0.3)
    if ctypes.windll.user32.GetForegroundWindow() != janela.handle:
        raise RuntimeError(
            "O Windows não trouxe o Chrome para frente. Clique na aba do Sicalc e tente novamente."
        )


def _url_atual(janela, win32clipboard, send_keys) -> str:
    from win32con import CF_UNICODETEXT

    janela.set_focus()
    send_keys("^l^c")
    time.sleep(0.2)
    try:
        win32clipboard.OpenClipboard()
        try:
            url = win32clipboard.GetClipboardData(CF_UNICODETEXT).strip()
        finally:
            win32clipboard.CloseClipboard()
    except Exception as exc:
        raise RuntimeError("Não foi possível ler a URL do Chrome pela barra de endereços.") from exc
    finally:
        send_keys("{ESC}")
    return url


def _codigo_ja_selecionado(janela, codigo: str) -> bool:
    for controle in _elementos(janela, ("Edit", "ComboBox")):
        if _id(controle) == "codReceitaPrincipal" or "codigo da receita" in _texto(_nome(controle)):
            return False
    esperado = re.sub(r"\D", "", codigo)
    for controle in _elementos(janela, ("ListItem",)):
        inicio = re.match(r"\s*(\d{4})\s*[-/]?\s*(\d{2})(?!\d)", _nome(controle))
        if inicio and "".join(inicio.groups()) == esperado:
            return True
    return False


def _validar_tela_calculo(janela, codigo: str, win32clipboard, send_keys) -> str:
    url = _url_atual(janela, win32clipboard, send_keys)
    partes = urlsplit(url)
    if (partes.scheme != "https" or partes.hostname != "sicalc.receita.fazenda.gov.br"
            or not partes.path.startswith("/sicalc/rapido/")
            or partes.path.endswith("/contribuinte")):
        raise RuntimeError(
            "Abra a tela do código após clicar em Continuar. "
            f"A aba ativa está em {partes.scheme}://{partes.hostname or '(sem domínio)'}{partes.path}."
        )
    if not _codigo_ja_selecionado(janela, codigo):
        _achar(janela, tipos=("Edit", "ComboBox"), ids=("codReceitaPrincipal",),
               nomes=("codigo da receita", "codigo de receita"), descricao="o campo do código")
    return url


def _selecionar_codigo(janela, codigo: str, send_keys) -> None:
    campo = _achar(janela, tipos=("Edit", "ComboBox"), ids=("codReceitaPrincipal",),
                  nomes=("codigo da receita", "codigo de receita"), descricao="o campo do código")
    _digitar(campo, codigo, send_keys, sair=False)
    esperado = re.sub(r"\D", "", codigo)
    limite = time.monotonic() + 15
    while time.monotonic() < limite:
        opcoes = []
        for controle in _elementos(janela, ("ListItem", "MenuItem", "Text")):
            nome = _nome(controle)
            inicio = re.match(r"\s*(\d{4})\s*[-/]?\s*(\d{2})(?!\d)", nome)
            if inicio and "".join(inicio.groups()) == esperado:
                opcoes.append(controle)
        for tipo in ("ListItem", "MenuItem", "Text"):
            grupo = [c for c in opcoes if c.element_info.control_type == tipo]
            if len(grupo) == 1:
                grupo[0].click_input()
                return
        time.sleep(0.3)
    raise RuntimeError(
        f"O código exato {codigo} não apareceu na lista acessível do Chrome. "
        f"Controles visíveis: {_resumo_controles(janela)}"
    )


def _esperar_campo(janela, ids: tuple[str, ...], nomes: tuple[str, ...], descricao: str,
                  timeout: float = 15):
    limite = time.monotonic() + timeout
    while time.monotonic() < limite:
        try:
            return _achar(janela, tipos=("Edit", "ComboBox"), ids=ids,
                          nomes=nomes, descricao=descricao)
        except RuntimeError:
            time.sleep(0.3)
    return _achar(janela, tipos=("Edit", "ComboBox"), ids=ids,
                 nomes=nomes, descricao=descricao)


def _clicar_botao(janela, identificador: str, nome: str, descricao: str) -> None:
    botao = _achar(janela, tipos=("Button",), ids=(identificador,),
                   nomes=(nome,), descricao=descricao)
    if not botao.is_enabled():
        raise RuntimeError(f"{descricao} está desabilitado na página.")
    botao.click_input()


def _selecionar_resultado(janela) -> None:
    limite = time.monotonic() + 30
    while time.monotonic() < limite:
        caixas = _elementos(janela, ("CheckBox",))
        if len(caixas) == 1:
            caixa = caixas[0]
            try:
                estado = caixa.get_toggle_state()
            except Exception as exc:
                raise RuntimeError("Não foi possível ler a caixa de seleção do DARF.") from exc
            if estado == 0:
                caixa.click_input()
            if caixa.get_toggle_state() != 1:
                raise RuntimeError("A caixa de seleção do DARF não ficou marcada.")
            return
        if len(caixas) > 1:
            raise RuntimeError(
                "Há mais de uma caixa de seleção visível; não é seguro escolher uma sem identificação. "
                f"Controles: {_resumo_controles(janela)}"
            )
        time.sleep(0.3)
    raise RuntimeError("O Sicalc não exibiu uma caixa de seleção para o resultado calculado.")


def _pdfs_novos(downloads: Path, anteriores: dict[Path, int]) -> list[Path]:
    if not downloads.is_dir():
        return []
    novos = []
    for arquivo in downloads.glob("*.pdf"):
        try:
            if arquivo in anteriores and arquivo.stat().st_mtime_ns <= anteriores[arquivo]:
                continue
            with arquivo.open("rb") as leitor:
                if leitor.read(4) == b"%PDF":
                    novos.append(arquivo)
        except OSError:
            pass
    return novos


def _salvar_dialogo(destino: Path, Desktop, send_keys) -> bool:
    desktop = Desktop(backend="uia")
    limite = time.monotonic() + 10
    while time.monotonic() < limite:
        dialogos = [j for j in desktop.windows() if _texto(j.window_text()) in ("salvar como", "save as")]
        if len(dialogos) == 1:
            dialogo = dialogos[0]
            campos = [c for c in dialogo.descendants(control_type="Edit") if _visivel(c)]
            por_nome = [c for c in campos if any(
                termo in _texto(_nome(c)) for termo in ("nome do arquivo", "file name")
            )]
            if len(por_nome) == 1:
                campo = por_nome[0]
            elif len(campos) == 1:
                campo = campos[0]
            else:
                raise RuntimeError("A janela Salvar como abriu, mas o campo do nome não foi identificado.")
            _digitar(campo, str(destino), send_keys, sair=False)
            send_keys("{ENTER}")
            return True
        time.sleep(0.2)
    return False


def _receber_pdf(janela, url_calculo: str, downloads: Path, destino: Path,
                 psutil, Desktop, win32clipboard, send_keys) -> Path:
    anteriores = {a: a.stat().st_mtime_ns for a in downloads.glob("*.pdf")} if downloads.is_dir() else {}
    _clicar_botao(janela, "btnDarf", "emitir darf", "o botão Emitir Darf")
    limite = time.monotonic() + 25
    url_pdf = ""
    visto_pdf_em = None
    while time.monotonic() < limite:
        novos = _pdfs_novos(downloads, anteriores)
        if len(novos) == 1:
            shutil.copy2(novos[0], destino)
            break
        if len(novos) > 1:
            raise RuntimeError("Mais de um PDF apareceu na pasta de downloads; não foi possível identificar o DARF.")
        try:
            janela_atual = _janela_ativa(psutil, Desktop)
            atual = _url_atual(janela_atual, win32clipboard, send_keys)
        except RuntimeError:
            atual = ""
        if atual and atual != url_calculo and (
            atual.startswith("blob:https://sicalc.receita.fazenda.gov.br/")
            or urlsplit(atual).hostname == "sicalc.receita.fazenda.gov.br"
        ):
            url_pdf = atual
            if visto_pdf_em is None:
                visto_pdf_em = time.monotonic()
            if time.monotonic() - visto_pdf_em >= 5:
                break
        time.sleep(0.5)
    if not destino.exists() and url_pdf:
        send_keys("^s")
        _salvar_dialogo(destino, Desktop, send_keys)
        limite = time.monotonic() + 30
        while not destino.is_file() and time.monotonic() < limite:
            novos = _pdfs_novos(downloads, anteriores)
            if len(novos) == 1:
                shutil.copy2(novos[0], destino)
                break
            if len(novos) > 1:
                raise RuntimeError("Mais de um PDF apareceu na pasta de downloads; não foi possível identificar o DARF.")
            time.sleep(0.3)
    if not destino.is_file():
        raise RuntimeError(
            "O PDF não foi encontrado. Verifique se o Chrome abriu uma nova aba ou salvou "
            "o arquivo em outra pasta; informe-a com --downloads."
        )
    with destino.open("rb") as leitor:
        assinatura = leitor.read(4)
    if assinatura != b"%PDF":
        destino.unlink(missing_ok=True)
        raise RuntimeError("O arquivo salvo não é um PDF; a emissão não foi confirmada.")
    return destino


def continuar_no_chrome_comum(*, codigo: str, valor: Decimal, periodo: str, cnpj: str,
                             downloads: Path, destino: Path, inspecionar: bool = False,
                             progresso: Callable[[str], None] = print) -> Path | None:
    """Assume a aba ativa do Chrome após o usuário concluir Contribuinte."""
    psutil, win32clipboard, Desktop, send_keys = _bibliotecas()
    if inspecionar:
        print("Controles visíveis da janela ativa:")
        janela = _janela_ativa(psutil, Desktop)
        print(_resumo_controles(janela))
        return None
    janela = _janela_ativa(psutil, Desktop)
    url_calculo = _validar_tela_calculo(janela, codigo, win32clipboard, send_keys)
    campos_cnpj = [c for c in _elementos(janela, ("Edit",)) if _id(c) == "cnpjFormatado"]
    if len(campos_cnpj) == 1 and re.sub(r"\D", "", _valor(campos_cnpj[0])) != cnpj:
        raise RuntimeError("O CNPJ exibido no Sicalc difere do informado na planilha.")
    progresso("Tela do código confirmada. Preenchendo os dados...")
    if not _codigo_ja_selecionado(janela, codigo):
        _selecionar_codigo(janela, codigo, send_keys)
    progresso("Código da receita selecionado.")
    _esperar_campo(janela, ("valorPrincipal", "valorPrincipalImposto"),
                  ("valor principal",), "o valor principal")
    # A troca de código atualiza os campos assincronamente e pode apagar a
    # primeira digitação feita logo após a escolha na lista.
    time.sleep(0.5)
    for _ in range(3):
        campo_periodo = _esperar_campo(janela, ("dataPA",),
                                      ("periodo de apuracao",), "o período de apuração")
        _digitar(campo_periodo, periodo, send_keys)
        time.sleep(0.3)
        if re.sub(r"\D", "", _valor(campo_periodo)) == re.sub(r"\D", "", periodo):
            break
    else:
        raise RuntimeError("O período exibido no Sicalc difere do informado.")
    vencimento = _esperar_campo(janela, ("dataVencimentoImposto",),
                               ("data de vencimento", "vencimento"), "a data de vencimento")
    limite = time.monotonic() + 15
    while not _valor(vencimento) and time.monotonic() < limite:
        time.sleep(0.3)
    if not _valor(vencimento):
        raise RuntimeError("O Sicalc não preencheu a data de vencimento após o período.")
    progresso("Período e vencimento conferidos.")
    principal = _esperar_campo(janela, ("valorPrincipal", "valorPrincipalImposto"),
                              ("valor principal", "valor do principal"), "o valor principal")
    _digitar(principal, f"{valor:.2f}".replace(".", ","), send_keys)
    try:
        conferido = Decimal(_valor(principal).replace(".", "").replace(",", "."))
    except InvalidOperation as exc:
        raise RuntimeError("O valor principal não pôde ser lido após o preenchimento.") from exc
    if conferido != valor:
        raise RuntimeError("O valor principal exibido no Sicalc difere do informado.")
    progresso("Valor principal conferido. Calculando...")
    _clicar_botao(janela, "btnCalcular", "calcular", "o botão Calcular")
    _selecionar_resultado(janela)
    progresso("Resultado selecionado. Emitindo o PDF...")
    return _receber_pdf(janela, url_calculo, downloads, destino,
                        psutil, Desktop, win32clipboard, send_keys)

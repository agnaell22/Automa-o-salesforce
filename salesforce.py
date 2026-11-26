import os
import re
import shutil
import time
import zipfile
from collections.abc import Callable
from dataclasses import dataclass, field
from pathlib import Path

from selenium import webdriver
from selenium.common.exceptions import (
    ElementClickInterceptedException,
    ElementNotInteractableException,
    StaleElementReferenceException,
    TimeoutException,
    WebDriverException,
)
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support.ui import WebDriverWait


LOGIN_TIMEOUT_SECONDS = 300
PAGE_TIMEOUT_SECONDS = 30
SEARCH_RESULTS_TIMEOUT_SECONDS = 12
DOWNLOAD_TIMEOUT_SECONDS = 60
SEARCH_FIELD_STABLE_POLLS = 3
PHOTO_EXTENSION = re.compile(r"\.(?:jpe?g|png|gif|webp|heic|bmp)$", re.IGNORECASE)


class _SearchStaleError(RuntimeError):
    pass


@dataclass
class ResultadoNota:
    nota: str
    status: str = "Pendente"
    status_wfm: str = ""
    motivo: str = ""
    arquivos: list[str] = field(default_factory=list)
    erro: str = ""


def criar_navegador(download_dir: Path, profile_dir: Path | None) -> webdriver.Chrome:
    local_app_data = os.environ.get("LOCALAPPDATA")
    if not local_app_data:
        raise RuntimeError("A variável LOCALAPPDATA não está definida.")

    user_data_dir = profile_dir or (
        Path(local_app_data) / "SalesforceAutomation" / "ChromeProfile"
    )
    user_data_dir = user_data_dir.expanduser().resolve()
    user_data_dir.mkdir(parents=True, exist_ok=True)
    download_dir.mkdir(parents=True, exist_ok=True)
    options = Options()
    options.add_argument(f"--user-data-dir={user_data_dir}")
    options.page_load_strategy = "eager"
    options.add_experimental_option("detach", True)
    options.add_experimental_option(
        "prefs",
        {
            "download.default_directory": str(download_dir.resolve()),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
        },
    )
    try:
        driver = webdriver.Chrome(options=options)
    except WebDriverException as error:
        raise RuntimeError(
            "Não foi possível iniciar o Chrome. Feche todas as janelas do Chrome "
            "de automação e tente novamente."
        ) from error
    driver.maximize_window()
    driver.execute_cdp_cmd(
        "Page.setDownloadBehavior",
        {"behavior": "allow", "downloadPath": str(download_dir.resolve())},
    )
    return driver


def _wait_for_login(driver: webdriver.Chrome, base_url: str) -> None:
    driver.get(base_url)
    try:
        WebDriverWait(driver, LOGIN_TIMEOUT_SECONDS).until(
            lambda browser: "lightning.force.com" in browser.current_url
            and _visible_search_button(browser)
        )
    except TimeoutException as error:
        raise RuntimeError(
            "A sessão Salesforce não ficou pronta. Faça login manualmente no Chrome "
            "e execute o programa novamente. A janela do Chrome permanecerá aberta."
        ) from error
    except WebDriverException as error:
        raise RuntimeError(
            "A janela do Chrome foi fechada antes de o login terminar."
        ) from error


def _buscar_ordem(
    driver: webdriver.Chrome, wait: WebDriverWait, nota: str
) -> None:
    failed_steps: list[str] = []
    for attempt in range(2):
        current_step = ["preparar a busca"]
        try:
            _buscar_ordem_once(driver, wait, nota, current_step)
            return
        except StaleElementReferenceException as error:
            failed_steps.append(
                f"tentativa {attempt + 1}: {current_step[0]}"
            )
            if attempt:
                raise _SearchStaleError(
                    f"A busca da nota {nota} perdeu a referência à página do "
                    "Salesforce mesmo após tentar novamente. Etapas: "
                    f"{'; '.join(failed_steps)}. "
                    f"{error.msg.splitlines()[0]}"
                ) from error


def _fill_search_box(
    driver: webdriver.Chrome,
    search_box: WebElement,
    nota: str,
) -> None:
    try:
        if search_box.tag_name.casefold() == "lightning-input":
            search_box.send_keys(Keys.CONTROL, "a", Keys.BACKSPACE)
        else:
            search_box.clear()
        search_box.send_keys(nota)
    except ElementNotInteractableException as error:
        active_search_box = _active_search_box(driver)
        if not active_search_box:
            driver.execute_script("arguments[0].focus()", search_box)
            active_search_box = _active_search_box(driver) or search_box
        try:
            active_search_box.send_keys(Keys.CONTROL, "a", Keys.BACKSPACE)
            active_search_box.send_keys(nota)
        except StaleElementReferenceException:
            raise
        except WebDriverException:
            raise error


def _buscar_ordem_once(
    driver: webdriver.Chrome,
    wait: WebDriverWait,
    nota: str,
    current_step: list[str],
) -> None:
    current_step[0] = "verificar se há modal de arquivos aberto"
    _fechar_modal_arquivos(driver)
    current_step[0] = "localizar o campo de pesquisa"
    search_box = _wait_for_stable_search_box(driver, timeout=0.5)
    if not search_box:
        current_step[0] = "localizar e clicar no botão Pesquisar"
        try:
            search_button = wait.until(_visible_search_button)
            try:
                search_button.click()
            except ElementClickInterceptedException:
                search_box = _wait_for_stable_search_box(driver, timeout=1.5)
                if not search_box:
                    search_box = _active_search_box(driver)
                if not search_box:
                    raise
        except TimeoutException as error:
            raise RuntimeError(
                f"O botão de pesquisa do Salesforce não apareceu para a nota "
                f"{nota}."
            ) from error
        except StaleElementReferenceException:
            raise
        except WebDriverException as error:
            raise RuntimeError(
                f"Não foi possível abrir a busca do Salesforce: "
                f"{error.msg.splitlines()[0]}"
            ) from error

        if not search_box:
            current_step[0] = "aguardar o campo após abrir a pesquisa"
            search_box = _wait_for_stable_search_box(driver, timeout=1.5)
        if not search_box:
            search_box = _active_search_box(driver)
        if not search_box:
            current_step[0] = "reabrir a pesquisa pelo botão"
            try:
                search_button = wait.until(_visible_search_button)
                driver.execute_script("arguments[0].click()", search_button)
            except TimeoutException as error:
                raise RuntimeError(
                    f"O botão de pesquisa do Salesforce não apareceu para a "
                    f"nota {nota}."
                ) from error
            except StaleElementReferenceException:
                raise
            except WebDriverException as error:
                raise RuntimeError(
                    f"Não foi possível reabrir a busca do Salesforce: "
                    f"{error.msg.splitlines()[0]}"
                ) from error

            current_step[0] = "aguardar o campo após reabrir a pesquisa"
            search_box = _wait_for_stable_search_box(driver, timeout=5)
            if not search_box:
                search_box = _active_search_box(driver)

    if not search_box:
        raise RuntimeError(
            f"O campo de busca da nota {nota} não apareceu após abrir a pesquisa."
        )

    current_step[0] = "limpar e preencher o campo de pesquisa"
    try:
        _fill_search_box(driver, search_box, nota)
    except StaleElementReferenceException:
        raise
    except WebDriverException as error:
        raise RuntimeError(
            f"O campo de busca da nota {nota} não pôde ser focado, limpo ou preenchido: "
            f"{error.msg.splitlines()[0]}"
        ) from error

    previous_url = driver.current_url
    current_step[0] = "localizar e abrir o resultado da busca"
    try:
        result = WebDriverWait(driver, SEARCH_RESULTS_TIMEOUT_SECONDS).until(
            lambda browser: _work_order_search_result(browser, nota)
        )
    except TimeoutException as error:
        raise RuntimeError(
            f"A busca da nota {nota} não confirmou se existe uma Ordem de Trabalho; "
            "os resultados não carregaram dentro do tempo limite."
        ) from error

    if result == "not-found":
        raise LookupError(
            f"Nenhuma Ordem de Trabalho foi encontrada para a nota {nota}."
        )

    def work_order_opened(browser: webdriver.Chrome) -> bool:
        current_url = browser.current_url
        return (
            "/lightning/r/WorkOrder/" in current_url
            and current_url != previous_url
        )

    current_step[0] = "aguardar a abertura da Ordem de Serviço"
    try:
        WebDriverWait(driver, PAGE_TIMEOUT_SECONDS).until(work_order_opened)
    except TimeoutException as error:
        raise RuntimeError(
            f"A busca da nota {nota} não abriu uma nova Ordem de Trabalho; "
            "a página permaneceu no registro anterior."
        ) from error

def _visible_search_button(driver: webdriver.Chrome) -> WebElement | bool:
    return driver.execute_script(
        """
        function findVisibleButton(root) {
            const buttons = root.querySelectorAll(
                'button[aria-label="Pesquisar"], ' +
                'button[aria-label="Search Salesforce"], ' +
                'button.search-button, [role="button"][aria-label="Pesquisar"]'
            );
            for (const button of buttons) {
                const rect = button.getBoundingClientRect();
                const style = getComputedStyle(button);
                if (
                    !button.disabled &&
                    button.getAttribute('aria-disabled') !== 'true' &&
                    rect.width > 0 &&
                    rect.height > 0 &&
                    style.visibility !== 'hidden' &&
                    style.display !== 'none' &&
                    style.opacity !== '0'
                ) {
                    return button;
                }
            }
            return null;
        }

        function findSearchButton(root) {
            const button = findVisibleButton(root);
            if (button) {
                return button;
            }

            for (const element of root.querySelectorAll('*')) {
                if (element.shadowRoot) {
                    const button = findSearchButton(element.shadowRoot);
                    if (button) {
                        return button;
                    }
                }
            }
            return null;
        }
        return findSearchButton(document);
        """
    ) or False


def _active_search_box(driver: webdriver.Chrome) -> WebElement | bool:
    active_element = driver.switch_to.active_element
    tag_name = active_element.tag_name.casefold()
    if tag_name in {"input", "textarea"}:
        field_type = (active_element.get_attribute("type") or "").casefold()
        role = (active_element.get_attribute("role") or "").casefold()
        label = " ".join(
            (
                active_element.get_attribute("placeholder") or "",
                active_element.get_attribute("aria-label") or "",
            )
        ).casefold()
        if field_type == "search" or role == "searchbox" or any(
            term in label for term in ("pesquisar", "search")
        ):
            return active_element
    elif (
        tag_name == "lightning-input"
        and "sainput" in (active_element.get_attribute("class") or "").casefold()
    ):
        return active_element
    return False


def _wait_for_stable_search_box(
    driver: webdriver.Chrome,
    timeout: float,
) -> WebElement | bool:
    previous_id: str | None = None
    previous_rect: tuple[int, ...] | None = None
    stable_polls = 0

    def stable_search_box(browser: webdriver.Chrome) -> WebElement | bool:
        nonlocal previous_id, previous_rect, stable_polls
        try:
            search_box = _visible_search_box(browser)
            if not search_box:
                previous_id = None
                previous_rect = None
                stable_polls = 0
                return False

            rect = browser.execute_script(
                "const r=arguments[0].getBoundingClientRect(); "
                "return [r.x, r.y, r.width, r.height]",
                search_box,
            )
            if not rect or rect[2] <= 0 or rect[3] <= 0:
                previous_id = None
                previous_rect = None
                stable_polls = 0
                return False

            current_rect = tuple(round(value) for value in rect)
            current_id = search_box.id
            if current_id == previous_id and current_rect == previous_rect:
                stable_polls += 1
            else:
                previous_id = current_id
                previous_rect = current_rect
                stable_polls = 1

            return search_box if stable_polls >= SEARCH_FIELD_STABLE_POLLS else False
        except StaleElementReferenceException:
            previous_id = None
            previous_rect = None
            stable_polls = 0
            return False

    try:
        return WebDriverWait(driver, timeout, poll_frequency=0.25).until(
            stable_search_box
        )
    except TimeoutException:
        return False


def _visible_search_box(driver: webdriver.Chrome) -> WebElement | bool:
    return driver.execute_script(
        """
        function isAccessibleFromPage(element) {
            const rect = element.getBoundingClientRect();
            let topElement = document.elementFromPoint(
                rect.left + rect.width / 2,
                rect.top + rect.height / 2
            );
            while (topElement && topElement.shadowRoot) {
                const nestedElement = topElement.shadowRoot.elementFromPoint(
                    rect.left + rect.width / 2,
                    rect.top + rect.height / 2
                );
                if (!nestedElement || nestedElement === topElement) {
                    break;
                }
                topElement = nestedElement;
            }
            return topElement === element;
        }

        function findSearchInput(root) {
            const candidates = root.querySelectorAll(
                'input[type="search"], input[role="searchbox"], ' +
                'input[placeholder*="pesquisar" i], input[aria-label*="pesquisar" i], ' +
                'input[placeholder*="search" i], input[aria-label*="search" i], ' +
                'lightning-input.saInput'
            );
            for (const element of candidates) {
                if (element.matches('input')) {
                    const role = element.getAttribute('role') || '';
                    const type = (element.getAttribute('type') || '').toLowerCase();
                    const label = [
                        element.getAttribute('placeholder') || '',
                        element.getAttribute('aria-label') || ''
                    ].join(' ').toLowerCase();
                    const isObjectTypePicker =
                        label.includes('por tipo de objeto') ||
                        label.includes('by object type');
                    const isSearchInput =
                        type === 'search' ||
                        role === 'searchbox' ||
                        (/pesquisar|search/.test(label) && role !== 'combobox');
                    const rect = element.getBoundingClientRect();
                    const style = getComputedStyle(element);
                    if (
                        isSearchInput &&
                        !isObjectTypePicker &&
                        !element.disabled &&
                        !element.readOnly &&
                        rect.width > 0 &&
                        rect.height > 0 &&
                        style.visibility !== 'hidden' &&
                        style.display !== 'none' &&
                        style.opacity !== '0' &&
                        style.pointerEvents !== 'none' &&
                        isAccessibleFromPage(element)
                    ) {
                        return element;
                    }
                }
            }

            for (const element of root.querySelectorAll('*')) {
                if (element.shadowRoot) {
                    const nestedInput = findSearchInput(element.shadowRoot);
                    if (nestedInput) {
                        return nestedInput;
                    }
                }
            }
            return null;
        }
        return findSearchInput(document);
        """
    ) or False


def _matching_work_order_result(
    driver: webdriver.Chrome, nota: str
) -> bool:
    return driver.execute_script(
        """
        function visible(element) {
            const rect = element.getBoundingClientRect();
            const style = getComputedStyle(element);
            return rect.width > 0 && rect.height > 0 &&
                style.visibility !== 'hidden' && style.display !== 'none' &&
                style.opacity !== '0';
        }

        function findWorkOrder(root, note) {
            const notePattern = new RegExp('(^|\\\\D)' + note + '(\\\\D|$)');
            for (const option of root.querySelectorAll('[role="option"]')) {
                const label = [
                    option.innerText || '',
                    option.textContent || '',
                    option.getAttribute('aria-label') || '',
                    option.getAttribute('title') || ''
                ].join(' ').replace(/\\s+/g, ' ').trim().toLowerCase();
                if (
                    visible(option) &&
                    !option.getAttribute('aria-disabled')?.match(/^true$/i) &&
                    notePattern.test(label) &&
                    /ordem\\s+(de\\s+)?(trabalho|servi[cç]o)|work\\s*order/i.test(label)
                ) {
                    return option;
                }
            }

            for (const element of root.querySelectorAll('*')) {
                if (element.shadowRoot) {
                    const option = findWorkOrder(element.shadowRoot, note);
                    if (option) {
                        return option;
                    }
                }
            }
            return null;
        }
        const result = findWorkOrder(document, arguments[0].toLowerCase());
        if (!result) {
            return false;
        }
        result.scrollIntoView({block: 'center'});
        result.click();
        return true;
        """,
        nota,
    )


def _no_work_order_results_visible(driver: webdriver.Chrome) -> bool:
    return bool(
        driver.execute_script(
            """
            const phrases = [
                'nenhum resultado',
                'nenhum registro',
                'no results',
                'no records'
            ];
            function hasNoResults(root) {
                for (const element of root.querySelectorAll(
                    '[role="status"], [role="alert"], [aria-live]'
                )) {
                    const text = (element.innerText || element.textContent || '')
                        .replace(/\\s+/g, ' ').toLowerCase();
                    if (phrases.some(phrase => text.includes(phrase))) {
                        return true;
                    }
                }
                for (const element of root.querySelectorAll('*')) {
                    if (
                        element.shadowRoot &&
                        hasNoResults(element.shadowRoot)
                    ) {
                        return true;
                    }
                }
                return false;
            }
            return hasNoResults(document);
            """
        )
    )


def _work_order_search_result(
    driver: webdriver.Chrome, nota: str
) -> str | bool:
    result = _matching_work_order_result(driver, nota)
    if result:
        return "opened"
    if _no_work_order_results_visible(driver):
        return "not-found"
    return False


def _abrir_aba_relacionados(
    driver: webdriver.Chrome,
    wait: WebDriverWait,
    nota: str,
) -> None:
    def related_tab_state(browser: webdriver.Chrome, activate: bool) -> str | bool:
        try:
            return browser.execute_script(
                """
                function visible(element) {
                    if (!element.isConnected) {
                        return false;
                    }
                    const rect = element.getBoundingClientRect();
                    const style = getComputedStyle(element);
                    return rect.width > 0 && rect.height > 0 &&
                        style.visibility !== 'hidden' &&
                        style.display !== 'none' && style.opacity !== '0';
                }

                function elementsIn(root, selector, result = []) {
                    result.push(...root.querySelectorAll(selector));
                    for (const element of root.querySelectorAll('*')) {
                        if (element.shadowRoot) {
                            elementsIn(element.shadowRoot, selector, result);
                        }
                    }
                    return result;
                }

                function hasVisibleAppointmentsCard(root) {
                    for (const article of elementsIn(
                        root,
                        'article, lightning-card'
                    )) {
                        if (!visible(article)) {
                            continue;
                        }
                        const text = (
                            article.innerText || article.textContent || ''
                        ).replace(/\\s+/g, ' ').toLowerCase();
                        if (text.includes('compromissos de serviços')) {
                            return true;
                        }
                    }
                    return false;
                }

                if (hasVisibleAppointmentsCard(document)) {
                    return 'content-ready';
                }

                const tabs = elementsIn(document, '[role="tab"]').filter(tab => {
                    const label = [
                        tab.innerText || '',
                        tab.textContent || '',
                        tab.getAttribute('aria-label') || '',
                        tab.getAttribute('title') || ''
                    ].join(' ').toLowerCase();
                    return visible(tab) &&
                        (label.includes('relacionado') || label.includes('related'));
                });
                const selected = tabs.find(
                    tab => tab.getAttribute('aria-selected') === 'true'
                );
                if (selected) {
                    return 'selected';
                }
                if (!arguments[0] || !tabs.length) {
                    return false;
                }

                const tab = tabs[0];
                tab.scrollIntoView({block: 'center'});
                tab.click();
                return 'clicked';
                """,
                activate,
            )
        except StaleElementReferenceException:
            return False

    try:
        wait.until(lambda browser: related_tab_state(browser, activate=True))
    except TimeoutException as error:
        raise RuntimeError(
            f"A aba Relacionados da Ordem de Serviço da nota {nota} não apareceu."
        ) from error

    try:
        wait.until(
            lambda browser: related_tab_state(browser, activate=False)
            in {"selected", "content-ready"}
        )
    except TimeoutException as error:
        raise RuntimeError(
            f"A aba Relacionados da Ordem de Serviço da nota {nota} "
            "não terminou de carregar."
        ) from error


def _compromissos_relacionados(
    wait: WebDriverWait,
    nota: str,
) -> list[str]:
    def related_appointments(browser: webdriver.Chrome):
        try:
            return browser.execute_script(
                """
                function visible(element) {
                    const rect = element.getBoundingClientRect();
                    const style = getComputedStyle(element);
                    return element.isConnected &&
                        rect.width > 0 && rect.height > 0 &&
                        style.visibility !== 'hidden' &&
                        style.display !== 'none' && style.opacity !== '0';
                }

                function elementsIn(root, selector, result = []) {
                    result.push(...root.querySelectorAll(selector));
                    for (const element of root.querySelectorAll('*')) {
                        if (element.shadowRoot) {
                            elementsIn(element.shadowRoot, selector, result);
                        }
                    }
                    return result;
                }

                const cards = elementsIn(
                    document,
                    'article, lightning-card'
                ).filter(article =>
                    visible(article) &&
                    (article.innerText || article.textContent || '')
                        .replace(/\\s+/g, ' ').toLowerCase()
                        .includes('compromissos de serviços')
                );
                for (const card of cards) {
                    const hrefs = [...new Set(
                        elementsIn(card, 'a')
                            .map(anchor => anchor.href)
                            .filter(href =>
                                /\\/lightning\\/r\\/(?:ServiceAppointment\\/)?08p[^/]+/i
                                    .test(href)
                            )
                    )];
                    if (hrefs.length) {
                        return hrefs;
                    }

                    const heading = elementsIn(card, 'h2').find(element =>
                        (element.innerText || element.textContent || '')
                            .includes('Compromissos de serviços')
                    );
                    const count = heading &&
                        (heading.innerText || heading.textContent || '')
                            .match(/\\((\\d+)\\)/);
                    if (count && count[1] === '0') {
                        return [];
                    }
                }
                return false;
                """
            )
        except StaleElementReferenceException:
            return False

    try:
        return wait.until(related_appointments)
    except TimeoutException as error:
        raise RuntimeError(
            f"A lista de Compromissos de Serviço da nota {nota} "
            "não carregou ou não contém registros identificáveis."
        ) from error


def _abrir_compromisso_servico(
    driver: webdriver.Chrome,
    wait: WebDriverWait,
    nota: str,
    appointment_url: str,
) -> None:
    appointment_id_match = re.search(
        r"/lightning/r/(?:ServiceAppointment/)?(08p[^/?]+)(?:/|$)",
        appointment_url,
    )
    if not appointment_id_match:
        raise RuntimeError(
            f"O link de Compromisso de Serviço da nota {nota} é inválido."
        )
    appointment_id = appointment_id_match.group(1)

    try:
        driver.get(appointment_url)
        wait.until(
            lambda browser: (
                "/lightning/r/" in browser.current_url
                and appointment_id in browser.current_url
            )
        )
    except TimeoutException as error:
        raise RuntimeError(
            f"Não foi possível abrir o Compromisso de Serviço {appointment_id} "
            f"relacionado à nota {nota}."
        ) from error
    except WebDriverException as error:
        raise RuntimeError(
            f"Não foi possível abrir o Compromisso de Serviço da nota {nota}: "
            f"{error.msg.splitlines()[0]}"
        ) from error


def _ler_valor_campo(driver: webdriver.Chrome, label: str) -> str:
    try:
        elements = driver.find_elements(
            By.XPATH,
            f"//li[.//*[normalize-space(.)='{label}']]",
        )
    except (TypeError, WebDriverException):
        return ""
    if not isinstance(elements, (list, tuple)):
        return ""
    for element in elements:
        lines = [
            line.strip()
            for line in element.text.splitlines()
            if line.strip() and line.strip() != label
        ]
        values = [line for line in lines if not line.startswith("Editar ")]
        if values:
            return " ".join(values)

    try:
        label_elements = driver.find_elements(
            By.XPATH,
            f"//*[normalize-space(.)='{label}']",
        )
    except (TypeError, WebDriverException):
        return ""
    if not isinstance(label_elements, (list, tuple)):
        return ""

    for label_element in label_elements:
        try:
            parent = label_element.find_element(By.XPATH, "..")
            lines = [
                line.strip()
                for line in parent.text.splitlines()
                if line.strip() and line.strip() != label
            ]
            values = [
                line
                for line in lines
                if not line.startswith("Editar ")
                and not line.startswith("Mostrar ")
                and not line.startswith("Exibir ")
            ]
            if values:
                return " ".join(values)
        except (StaleElementReferenceException, WebDriverException):
            continue
    return ""


def _ler_motivo(driver: webdriver.Chrome) -> str:
    values = [
        _ler_valor_campo(driver, "Motivo impedimento"),
        _ler_valor_campo(driver, "Justificativa do impedimento"),
    ]
    cards = driver.find_elements(
        By.XPATH,
        "//article[.//h2[contains(normalize-space(.), 'Motivos de impedimentos')]]",
    )
    for card in cards:
        text = card.text.splitlines()
        values.extend(
            line.strip()
            for line in text
            if line.strip()
            and "Motivos de impedimentos" not in line
            and not re.fullmatch(r"\(\d+\)", line.strip())
            and not line.strip().startswith("Mostrar ações")
            and not line.strip().startswith("Exibir tudo")
        )
    return " | ".join(dict.fromkeys(value for value in values if value))


def _ler_status_wfm(driver: webdriver.Chrome) -> str:
    values: list[str] = []
    for label in ("Status", "Status WFM", "Situação"):
        value = _ler_valor_campo(driver, label)
        if value:
            values.append(value)
    return " | ".join(dict.fromkeys(values))


def _click_files_download_button(driver: webdriver.Chrome) -> bool:
    try:
        return bool(
            driver.execute_script(
                """
        function findFilesButton(root) {
            for (const button of root.querySelectorAll(
                'button, [role="button"], lightning-button'
            )) {
                const label = [
                    button.innerText || button.textContent || '',
                    button.getAttribute('aria-label') || '',
                    button.getAttribute('title') || '',
                    button.getAttribute('label') || ''
                ].join(' ').replace(/\\s+/g, ' ').trim().toLowerCase();
                const rect = button.getBoundingClientRect();
                const style = getComputedStyle(button);
                if (
                    label.includes('baixar arquivos') &&
                    !button.disabled &&
                    button.getAttribute('aria-disabled') !== 'true' &&
                    rect.width > 0 &&
                    rect.height > 0 &&
                    style.visibility !== 'hidden' &&
                    style.display !== 'none' &&
                    style.opacity !== '0'
                ) {
                    button.scrollIntoView({block: 'center'});
                    button.click();
                    return true;
                }
            }

            for (const element of root.querySelectorAll('*')) {
                if (element.shadowRoot) {
                    const button = findFilesButton(element.shadowRoot);
                    if (button) {
                        return button;
                    }
                }
            }
            return false;
        }
        return findFilesButton(document);
        """
            )
        )
    except StaleElementReferenceException:
        return False


def _visible_file_modal(driver: webdriver.Chrome) -> bool:
    return bool(
        driver.execute_script(
            """
            function visible(element) {
                const rect = element.getBoundingClientRect();
                const style = getComputedStyle(element);
                return element.isConnected &&
                    rect.width > 0 && rect.height > 0 &&
                    style.visibility !== 'hidden' &&
                    style.display !== 'none' && style.opacity !== '0';
            }

            function elementsIn(root, selector) {
                const found = [...root.querySelectorAll(selector)];
                for (const element of root.querySelectorAll('*')) {
                    if (element.shadowRoot) {
                        found.push(...elementsIn(element.shadowRoot, selector));
                    }
                }
                return found;
            }

            return elementsIn(
                document,
                '[role="dialog"], .slds-modal__container'
            ).some(modal => {
                const text = (modal.innerText || modal.textContent || '')
                    .replace(/\\s+/g, ' ').toLowerCase();
                return visible(modal) && /(arquivo|files?)/i.test(text);
            });
            """
        )
    )


def _select_and_download_modal_images(
    driver: webdriver.Chrome,
) -> dict[str, bool] | bool:
    try:
        return (
            driver.execute_script(
                """
                const photoName = /\\.(?:jpe?g|png|gif|webp|heic|bmp)$/i;

                function visible(element) {
                    const rect = element.getBoundingClientRect();
                    const style = getComputedStyle(element);
                    return element.isConnected &&
                        rect.width > 0 && rect.height > 0 &&
                        style.visibility !== 'hidden' &&
                        style.display !== 'none' && style.opacity !== '0';
                }

                function elementsIn(root, selector) {
                    const found = [...root.querySelectorAll(selector)];
                    for (const element of root.querySelectorAll('*')) {
                        if (element.shadowRoot) {
                            found.push(...elementsIn(element.shadowRoot, selector));
                        }
                    }
                    return found;
                }

                function findModal() {
                    return elementsIn(
                        document,
                        '[role="dialog"], .slds-modal__container'
                    ).reverse().find(modal => {
                        const text = (modal.innerText || modal.textContent || '')
                            .replace(/\\s+/g, ' ').toLowerCase();
                        return visible(modal) &&
                            /(arquivo|files?)/i.test(text);
                    });
                }

                function photoRows(modal) {
                    return elementsIn(modal, '[role="row"]')
                        .filter(visible)
                        .map(row => {
                            const text = row.innerText || row.textContent || '';
                            const filename = text
                                .split(/\\r?\\n/)
                                .map(line => line.trim())
                                .find(line => photoName.test(line));
                            const checkbox = elementsIn(
                                row,
                                'input[type="checkbox"]'
                            ).find(input => !input.disabled);
                            return filename
                                ? {filename, checkbox}
                                : null;
                        })
                        .filter(Boolean);
                }

                let modal = findModal();
                if (!modal) {
                    return false;
                }

                const modalText = (
                    modal.innerText || modal.textContent || ''
                ).replace(/\\s+/g, ' ').toLowerCase();
                let rows = photoRows(modal);
                if (!rows.length) {
                    const hasFileRows = elementsIn(modal, '[role="row"]')
                        .filter(visible)
                        .some(row => {
                            const text = row.innerText || row.textContent || '';
                            return text.split(/\\r?\\n/).some(line =>
                                /\\.[a-z0-9]{2,5}$/i.test(line.trim())
                            );
                        });
                    return (
                        /nenhum arquivo|sem arquivos|no files/.test(modalText) ||
                        hasFileRows
                    )
                        ? {empty: true}
                        : false;
                }
                if (rows.some(row => !row.checkbox)) {
                    return false;
                }

                const imageCount = rows.length;
                for (let index = 0; index < imageCount; index += 1) {
                    modal = findModal();
                    if (!modal) {
                        return false;
                    }
                    rows = photoRows(modal);
                    if (rows.length !== imageCount) {
                        return false;
                    }
                    if (!rows[index].checkbox) {
                        return false;
                    }
                    if (!rows[index].checkbox.checked) {
                        rows[index].checkbox.click();
                    }
                }

                modal = findModal();
                if (!modal) {
                    return false;
                }
                rows = photoRows(modal);
                if (
                    rows.length !== imageCount ||
                    rows.some(row => !row.checkbox) ||
                    rows.some(row => !row.checkbox.checked)
                ) {
                    return false;
                }

                const downloadButton = elementsIn(
                    modal,
                    'button, [role="button"], lightning-button'
                ).find(button => {
                    const label = [
                        button.innerText || button.textContent || '',
                        button.getAttribute('aria-label') || '',
                        button.getAttribute('title') || '',
                        button.getAttribute('label') || ''
                    ].join(' ').replace(/\\s+/g, ' ').trim().toLowerCase();
                    return visible(button) &&
                        label === 'baixar' &&
                        !button.disabled &&
                        button.getAttribute('aria-disabled') !== 'true';
                });
                if (!downloadButton) {
                    return false;
                }

                downloadButton.click();
                return {started: true};
                """
            )
            or False
        )
    except StaleElementReferenceException:
        return False


def _baixar_arquivos_relacionados(
    driver: webdriver.Chrome,
    wait: WebDriverWait,
    nota: str,
    destination: Path,
) -> list[str]:
    result = _baixar_arquivos_relacionados_impl(
        driver,
        wait,
        nota,
        destination,
    )
    _fechar_modal_arquivos(driver)
    return result


def _baixar_arquivos_relacionados_impl(
    driver: webdriver.Chrome,
    wait: WebDriverWait,
    nota: str,
    destination: Path,
) -> list[str]:
    before = {path.name for path in destination.iterdir()}
    try:
        wait.until(_click_files_download_button)
    except TimeoutException as error:
        raise RuntimeError(
            f"O botão de arquivos não carregou para a nota {nota}."
        ) from error
    except WebDriverException as error:
        raise RuntimeError(
            f"Não foi possível iniciar os arquivos da nota {nota}: "
            f"{error.msg.splitlines()[0]}"
        ) from error

    def modal_or_download_started(browser: webdriver.Chrome) -> str | bool:
        try:
            if _visible_file_modal(browser):
                return "modal"

            new_files = [
                path
                for path in destination.iterdir()
                if path.is_file() and path.name not in before
            ]
            if new_files:
                return "direct"

            body_text = browser.find_element(By.TAG_NAME, "body").text.casefold()
            if any(
                phrase in body_text
                for phrase in (
                    "o download iniciará em breve",
                    "o download começará em breve",
                    "download will begin shortly",
                )
            ):
                return "direct"
            return False
        except StaleElementReferenceException:
            return False

    try:
        action_state = WebDriverWait(
            driver, PAGE_TIMEOUT_SECONDS
        ).until(modal_or_download_started)
    except TimeoutException as error:
        raise RuntimeError(
            f"A ação de baixar não abriu uma lista de arquivos nem iniciou "
            f"o download para a nota {nota}."
        ) from error

    if action_state == "modal":
        try:
            modal_result = WebDriverWait(
                driver, PAGE_TIMEOUT_SECONDS
            ).until(_select_and_download_modal_images)
        except TimeoutException as error:
            raise RuntimeError(
                f"Não foi possível selecionar e iniciar o download das imagens "
                f"da nota {nota} no modal de arquivos."
            ) from error
        if modal_result.get("empty"):
            return []
        if not modal_result.get("started"):
            raise RuntimeError(
                f"O modal não confirmou a seleção e o início do download "
                f"para a nota {nota}."
            )
    elif action_state != "direct":
        raise RuntimeError(
            f"A ação de download não foi confirmada para a nota {nota}."
        )
    deadline = time.monotonic() + DOWNLOAD_TIMEOUT_SECONDS
    downloaded: list[Path] = []
    last_snapshot: tuple[tuple[str, int], ...] | None = None
    stable_since: float | None = None
    while time.monotonic() < deadline:
        pending = [
            path
            for path in destination.glob("*.crdownload")
            if path.name not in before
        ]
        new_files = [
            path
            for path in destination.iterdir()
            if path.is_file()
            and path.name not in before
            and path.suffix.lower() != ".crdownload"
        ]
        snapshot = tuple(sorted((path.name, path.stat().st_size) for path in new_files))
        if new_files and not pending and snapshot == last_snapshot:
            if stable_since is not None and time.monotonic() - stable_since >= 2:
                downloaded = new_files
                break
        else:
            stable_since = time.monotonic() if new_files and not pending else None
            last_snapshot = snapshot if new_files and not pending else None
        time.sleep(0.5)

    if not downloaded:
        raise TimeoutError(
            f"Nenhum arquivo foi baixado para a nota {nota} dentro do tempo limite."
        )

    return _extract_downloaded_images(downloaded, nota, destination)


def _unique_image_path(destination: Path, nota: str, filename: str) -> Path:
    safe_name = Path(filename.replace("\\", "/")).name
    target = destination / f"{nota}_{safe_name}"
    suffix = 2
    while target.exists():
        target = destination / f"{nota}_{suffix}_{safe_name}"
        suffix += 1
    return target


def _extract_downloaded_images(
    downloaded: list[Path],
    nota: str,
    destination: Path,
) -> list[str]:
    image_paths: list[Path] = []
    for path in downloaded:
        if path.suffix.lower() == ".zip" or zipfile.is_zipfile(path):
            try:
                with zipfile.ZipFile(path) as archive:
                    images = [
                        member
                        for member in archive.infolist()
                        if not member.is_dir()
                        and PHOTO_EXTENSION.search(
                            Path(member.filename.replace("\\", "/")).name
                        )
                    ]
                    for member in images:
                        filename = Path(member.filename.replace("\\", "/")).name
                        target = _unique_image_path(destination, nota, filename)
                        try:
                            with archive.open(member) as source, target.open(
                                "wb"
                            ) as output:
                                shutil.copyfileobj(source, output)
                        except OSError:
                            target.unlink(missing_ok=True)
                            raise
                        image_paths.append(target)
            except (OSError, zipfile.BadZipFile) as error:
                raise RuntimeError(
                    f"Não foi possível extrair o arquivo ZIP '{path.name}': {error}"
                ) from error
            path.unlink()
        elif PHOTO_EXTENSION.search(path.name):
            target = _unique_image_path(destination, nota, path.name)
            path.replace(target)
            image_paths.append(target)

    return [path.name for path in image_paths]


def _fechar_modal_arquivos(driver: webdriver.Chrome) -> bool:
    close_result = driver.execute_script(
        """
        function visible(element) {
            const rect = element.getBoundingClientRect();
            const style = getComputedStyle(element);
            return element.isConnected &&
                rect.width > 0 && rect.height > 0 &&
                style.visibility !== 'hidden' && style.display !== 'none' &&
                style.opacity !== '0';
        }

        function elementsIn(root, selector) {
            const elements = [...root.querySelectorAll(selector)];
            for (const element of root.querySelectorAll('*')) {
                if (element.shadowRoot) {
                    elements.push(...elementsIn(element.shadowRoot, selector));
                }
            }
            return elements;
        }

        const modals = elementsIn(
            document,
            '[role="dialog"], .slds-modal__container'
        );
        for (const modal of modals) {
            if (!visible(modal)) {
                continue;
            }
            const text = (modal.innerText || modal.textContent || '')
                .replace(/\\s+/g, ' ').toLowerCase();
            if (!/(arquivo|files?)/i.test(text)) {
                continue;
            }

            const closeButton = elementsIn(
                modal,
                'button, [role="button"], lightning-button'
            ).find(button => {
                const label = [
                    button.innerText || button.textContent || '',
                    button.getAttribute('aria-label') || '',
                    button.getAttribute('title') || '',
                    button.getAttribute('label') || ''
                ].join(' ').replace(/\\s+/g, ' ').trim().toLowerCase();
                return visible(button) &&
                    !button.disabled &&
                    /cancelar e fechar|fechar|close/.test(label);
            });
            if (!closeButton) {
                return 'missing-close';
            }
            closeButton.click();
            return 'clicked';
        }
        return 'absent';
        """
    )
    if close_result == "absent":
        return False
    if close_result != "clicked":
        raise RuntimeError(
            "A janela de arquivos permaneceu aberta, mas não foi possível "
            "localizar o botão para fechá-la."
        )

    try:
        closed = WebDriverWait(driver, 5).until(
            lambda browser: not _visible_file_modal(browser)
        )
    except StaleElementReferenceException:
        closed = False
    if not closed:
        raise RuntimeError("A janela de arquivos não fechou após a consulta.")
    return True


def _fechar_aba_registro_ativa(driver: webdriver.Chrome) -> bool:
    tab, close_button = driver.execute_script(
        """
            const recordTab = /ordem de trabalho|work order|compromisso de servi[cç]o|service appointment/i;
            const closeTab = /^(fechar|close)(\\s|$)/i;
            const activeTab = [...document.querySelectorAll(
                '[role="tab"][aria-selected="true"]'
            )].find(tab => recordTab.test([
                tab.innerText || '',
                tab.textContent || '',
                tab.getAttribute('aria-label') || '',
                tab.getAttribute('title') || ''
            ].join(' ')));
            if (!activeTab) {
                return [null, null];
            }

            for (
                let container = activeTab, depth = 0;
                container && depth < 5;
                container = container.parentElement, depth += 1
            ) {
                const closeButton = [...container.querySelectorAll(
                    'button, [role="button"]'
                )].find(button => {
                    const label = [
                        button.getAttribute('aria-label') || '',
                        button.getAttribute('title') || '',
                        button.innerText || '',
                        button.textContent || ''
                    ].join(' ').replace(/\\s+/g, ' ').trim();
                    const rect = button.getBoundingClientRect();
                    const style = getComputedStyle(button);
                    return closeTab.test(label) &&
                        rect.width > 0 && rect.height > 0 &&
                        style.visibility !== 'hidden' &&
                        style.display !== 'none';
                });
                if (closeButton) {
                    return [activeTab, closeButton];
                }
            }

            return [activeTab, null];
            """
    )
    if not tab or not close_button:
        return False

    def tab_closed(browser: webdriver.Chrome) -> bool:
        try:
            return not browser.execute_script(
                "return arguments[0].isConnected",
                tab,
            )
        except StaleElementReferenceException:
            return True

    try:
        driver.execute_script("arguments[0].click()", close_button)
    except StaleElementReferenceException:
        pass

    WebDriverWait(driver, 5).until(tab_closed)
    return True


def _processar_nota(
    driver: webdriver.Chrome,
    output_dir: Path,
    nota: str,
) -> ResultadoNota:
    result = ResultadoNota(nota=nota)
    wait = WebDriverWait(driver, PAGE_TIMEOUT_SECONDS)
    stage = "pesquisar a nota"
    work_order_opened = False
    try:
        _buscar_ordem(driver, wait, nota)
        work_order_opened = True
        work_order_url = driver.current_url
        work_order_id = re.search(
            r"/lightning/r/WorkOrder/([^/?]+)",
            work_order_url,
        )
        if work_order_id:
            print(f"Nota {nota}: Ordem de Serviço {work_order_id.group(1)}")
        stage = "abrir Relacionados na Ordem de Serviço"
        _abrir_aba_relacionados(driver, wait, nota)
        stage = "localizar Compromissos de Serviço relacionados"
        appointments = _compromissos_relacionados(wait, nota)
        if not appointments:
            result.status = "Sem compromisso"
            return result

        note_dir = output_dir / nota
        note_dir.mkdir(parents=True, exist_ok=True)
        driver.execute_cdp_cmd(
            "Page.setDownloadBehavior",
            {"behavior": "allow", "downloadPath": str(note_dir.resolve())},
        )
        appointment_errors: list[str] = []
        for index, appointment_url in enumerate(appointments):
            try:
                appointment_id = re.search(
                    r"/lightning/r/(?:ServiceAppointment/)?(08p[^/?]+)",
                    appointment_url,
                )
                if appointment_id:
                    print(
                        f"Nota {nota}: Compromisso de Serviço "
                        f"{appointment_id.group(1)}"
                    )
                if index:
                    driver.get(work_order_url)
                    wait.until(
                        lambda browser: "/lightning/r/WorkOrder/"
                        in browser.current_url
                    )
                    _abrir_aba_relacionados(driver, wait, nota)

                stage = "abrir o Compromisso de Serviço relacionado"
                _abrir_compromisso_servico(
                    driver,
                    wait,
                    nota,
                    appointment_url,
                )
                stage = "ler o motivo de impedimento"
                result.motivo = " | ".join(
                    dict.fromkeys(
                        value
                        for value in (result.motivo, _ler_motivo(driver))
                        if value
                    )
                )
                stage = "ler o status WFM"
                appointment_status = _ler_status_wfm(driver)
                if not result.status_wfm:
                    result.status_wfm = appointment_status
                stage = "baixar os arquivos do Compromisso de Serviço"
                downloaded_files = _baixar_arquivos_relacionados(
                    driver,
                    wait,
                    nota,
                    note_dir,
                )
                result.arquivos.extend(downloaded_files)
                if downloaded_files:
                    print(
                        f"Nota {nota}: arquivos salvos em {note_dir}: "
                        f"{', '.join(downloaded_files)}"
                    )
            except (TimeoutError, TimeoutException, RuntimeError) as error:
                appointment_errors.append(str(error))
            except WebDriverException as error:
                appointment_errors.append(
                    f"Falha do navegador ao {stage}: "
                    f"{error.msg.splitlines()[0]}"
                )

        if appointment_errors:
            result.status = "Parcial" if result.arquivos else "Erro"
            result.erro = " | ".join(appointment_errors)
        else:
            result.status = "Concluída" if result.arquivos else "Sem imagens"
    except LookupError as error:
        result.status = "Não encontrada"
        result.erro = str(error)
    except _SearchStaleError:
        raise
    except (TimeoutError, TimeoutException, RuntimeError) as error:
        result.status = "Erro"
        result.erro = str(error)
    except WebDriverException as error:
        result.status = "Erro"
        result.erro = (
            f"Falha do navegador ao {stage}: {error.msg.splitlines()[0]}"
        )
    finally:
        if work_order_opened:
            try:
                if not _fechar_aba_registro_ativa(driver):
                    raise RuntimeError(
                        f"Não foi possível fechar a guia da Ordem de Serviço "
                        f"da nota {nota}."
                    )
            except (TimeoutError, TimeoutException, RuntimeError) as error:
                result.status = "Parcial" if result.arquivos else "Erro"
                result.erro = " | ".join(
                    filter(
                        None,
                        (
                            result.erro,
                            str(error),
                        ),
                    )
                )
            except WebDriverException as error:
                result.status = "Parcial" if result.arquivos else "Erro"
                result.erro = " | ".join(
                    filter(
                        None,
                        (
                            result.erro,
                            f"Falha ao fechar a guia da nota {nota}: "
                            f"{error.msg.splitlines()[0]}",
                        ),
                    )
                )
    return result


def consultar_notas(
    notes: list[str],
    salesforce_url: str,
    output_dir: Path,
    profile_dir: Path | None,
    on_result: Callable[[ResultadoNota], None] | None = None,
    output_fn: Callable[[str], None] = print,
) -> list[ResultadoNota]:
    startup_started = time.monotonic()
    output_fn("Iniciando o Chrome para consultar o Salesforce...")
    driver = criar_navegador(output_dir, profile_dir)
    output_fn(
        f"Chrome iniciado em {time.monotonic() - startup_started:.1f}s; "
        "conectando ao Salesforce..."
    )
    results: list[ResultadoNota] = []
    session_ready = False
    browser_session_lost = False
    try:
        login_started = time.monotonic()
        _wait_for_login(driver, salesforce_url)
        session_ready = True
        output_fn(
            f"Salesforce pronto em {time.monotonic() - login_started:.1f}s; "
            f"iniciando consulta de {len(notes)} nota(s)..."
        )
        for index, note in enumerate(notes):
            note_started = time.monotonic()
            if index == 0:
                output_fn(f"Iniciando a primeira busca: nota {note}...")
            search_unstable = False
            try:
                result = _processar_nota(driver, output_dir, note)
            except _SearchStaleError as error:
                result = ResultadoNota(
                    nota=note,
                    status="Erro",
                    erro=str(error),
                )
                search_unstable = True
            try:
                driver.current_url
            except WebDriverException as error:
                browser_session_lost = True
                result.status = "Erro"
                result.erro = " | ".join(
                    filter(
                        None,
                        (
                            result.erro,
                            "A sessão do Chrome/Selenium foi encerrada; "
                            "a fila de consultas foi interrompida.",
                            error.msg.splitlines()[0],
                        ),
                    )
                )
            results.append(result)
            if on_result is not None:
                on_result(result)
            output_fn(
                f"Nota {note}: {result.status}"
                + (f" ({result.erro})" if result.erro else "")
                + f" [{time.monotonic() - note_started:.1f}s]"
            )
            if browser_session_lost:
                output_fn(
                    "Chrome/Selenium não está mais disponível; interrompendo "
                    f"a fila após a nota {note}."
                )
                break
            if search_unstable:
                output_fn(
                    "A pesquisa do Salesforce continuou instável após a "
                    "tentativa de recuperação; fila interrompida. As notas "
                    "restantes não foram processadas."
                )
                break
    finally:
        if session_ready:
            try:
                driver.quit()
            except WebDriverException:
                driver.service.stop()
        else:
            driver.service.stop()
    return results

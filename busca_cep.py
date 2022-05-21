from playwright.sync_api import sync_playwright
import openpyxl

# Carregando arquivo
book = openpyxl.load_workbook('base_cep.xlsx')
# Selecionando uma página
planilha = book['Planilha1']
cep = []

for rows in planilha.iter_rows(min_row=2,max_row=26,min_col=1,max_col=1):
    cep.append(rows[0].value)


with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    page = browser.new_page()
    j = 2
    for i in cep: 

        page.goto("https://buscacepinter.correios.com.br/app/endereco/index.php")
        # page.wait_for_timeout(1000)
        # Colar o CEP name = "endereco"
        page.fill("input[name='endereco']",i)
        # page.wait_for_timeout(1000)
        # Clicar no botão buscar name="btn_pesquisar"
        page.click("button[name='btn_pesquisar']")
        # page.wait_for_timeout(1000)
        planilha['B' + str(j)] = page.text_content("td[data-th='Logradouro/Nome']")
        planilha['C' + str(j)] = page.text_content("td[data-th='Bairro/Distrito']")
        planilha['D' + str(j)] = page.text_content("td[data-th='Localidade/UF']")

        print(page.text_content("td[data-th='Logradouro/Nome']"))
        print(page.text_content("td[data-th='Bairro/Distrito']"))
        print(page.text_content("td[data-th='Localidade/UF']"))

        j = j + 1
book.save('base2.xlsx')

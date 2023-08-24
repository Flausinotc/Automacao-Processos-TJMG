# Documentação do Script de Automação para Extração de Informações de Processos Jurídicos

## Introdução

Este script em Python foi desenvolvido com o intuito de automatizar o processo de extração de informações de processos jurídicos a partir do site de consulta pública do Tribunal de Justiça de Minas Gerais (TJMG). O script utiliza as bibliotecas Selenium e openpyxl para interagir com o navegador e salvar os dados em um arquivo Excel.

## Créditos

Este script foi inspirado e adaptado a partir do trabalho original de Jhonatan de Souza do canal [DevAprender](https://www.youtube.com/c/DevAprender). Os créditos pelo desenvolvimento do conceito e a inspiração para este código são atribuídos a ele.

## Dependências

O script requer a instalação das seguintes bibliotecas:

- **Selenium:** Biblioteca para automação de testes e interações com navegadores web.
- **openpyxl:** Biblioteca para manipulação de arquivos Excel.

As dependências podem ser instaladas usando o gerenciador de pacotes `pip`. Para instalar as dependências, execute os seguintes comandos:

```
pip install selenium
pip install openpyxl
```

## Funcionalidades Principais

O script realiza as seguintes etapas:

1. Inicializa um driver do Chrome para automação de navegação.
2. Acessa o site de consulta pública do TJMG.
3. Insere o número da Ordem dos Advogados do Brasil (OAB) e seleciona o estado desejado.
4. Clica no botão de pesquisa e aguarda o carregamento dos resultados.
5. Itera sobre os links dos processos encontrados.
6. Para cada processo, clica no link para acessar a página de detalhes.
7. Extrai informações relevantes do processo, incluindo número, data de distribuição e movimentações.
8. Abre ou cria um arquivo Excel chamado 'dados.xlsx'.
9. Preenche as informações do processo na planilha Excel.
10. Salva as alterações no arquivo Excel.
11. Fecha a janela atual e retorna à página de resultados.
12. Repete o processo para todos os processos encontrados.
13. Fecha o driver do Chrome ao final da execução.

## Uso

Antes de executar o script, verifique os seguintes pontos:

- Garanta que o ChromeDriver esteja instalado e configurado corretamente.
- Verifique se os seletores XPath usados para localizar os elementos estão atualizados.
- Modifique o número da OAB (`numero_oba`) e outros parâmetros conforme necessário.

Execute o script utilizando o interpretador Python:

```
python nome_do_script.py
```

## Notas Finais

Este script foi desenvolvido com fins educacionais e para automação de tarefas específicas. Certifique-se de estar ciente dos termos de uso e ética ao realizar a automação de interações na web.

---

Certifique-se de verificar e ajustar todos os detalhes da documentação para se adequar às suas necessidades antes de utilizá-la.

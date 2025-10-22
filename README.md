#  Validador de Limites VARCHAR em Planilha Excel

Este script em Python verifica se os valores de uma planilha do Excel ultrapassam os limites definidos na primeira linha (ex: `VARCHAR(45)`).

---

##  Como funciona

- A **primeira linha** da planilha deve conter os limites de cada coluna no formato `VARCHAR(x)`.
- As **linhas seguintes** contêm os dados que serão verificados.
- O script informa, no terminal, quais valores excedem o limite permitido.

---

##  Tecnologias utilizadas

- Python
- openpyxl  
- re (expressões regulares)

---

##  Instalação

Clone o repositório e instale a biblioteca necessária:

```bash
git clone https://github.com/SEU-USUARIO/validador-varchar-excel.git
cd validador-varchar-excel
pip install openpyxl

# Certificado Fatec Rio Claro

Projeto desenvolvido com foco em automação de geração de certificados em PDF a partir de planilhas do Excel.

## Funcionalidade

Este projeto permite:
- Ler dados de uma planilha `alunos.xlsx`
- Gerar certificados individuais em PDF para cada aluno
- Personalizar o nome do curso, carga horária e responsável diretamente pela aba "Evento" da planilha

## Estrutura da planilha

O projeto usa a planilha `alunos.xlsx` com as seguintes abas:

### Aba `Evento`:
| Célula | Conteúdo               |
|--------|------------------------|
| B1     | Nome do curso          |
| B2     | Carga horária (em horas) |
| B3     | Nome do responsável    |

### Aba `certificados`:
| Coluna A       | Coluna B     |
|----------------|--------------|
| Nome do Aluno  | CPF do Aluno |

## Requisitos

- Python 3.8 ou superior
- `reportlab`
- `openpyxl`

## Instalação

### No ambiente Linux

```bash
git clone https://github.com/orlandosaraivajr/certificado_fatec_rio_claro.git
cd certificado_fatec_rio_claro/
virtualenv -p python3 venv
source venv/bin/activate
pip install -r requirements.txt
python main.py 
```

### No ambiente Windows

```bash
git clone https://github.com/orlandosaraivajr/certificado_fatec_rio_claro.git
cd certificado_fatec_rio_claro/
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

## Gerando certificados

## Personalização

Você pode adicionar a logomarca da instituição no arquivo `logo.png`. Ela será automaticamente incluída no topo de cada certificado.

## Licença

Este projeto é de uso educacional e livre para fins acadêmicos.

---

Desenvolvido por Orlando Saraiva Jr - Fatec Rio Claro 

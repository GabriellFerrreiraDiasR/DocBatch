# DocBatch

**Gerador em massa de documentos personalizados.**

O DocBatch cruza uma planilha Excel (`.xlsx`) com um modelo Word (`.docx`), gera um documento personalizado para cada linha e entrega todos os resultados em um único arquivo ZIP. Tudo roda localmente, sem depender de APIs externas ou serviços pagos.

> **Status:** em desenvolvimento (MVP). As instruções de execução abaixo descrevem o ambiente previsto e serão ajustadas conforme a implementação avança.

## Sumário

- [Problema e solução](#problema-e-solução)
- [Objetivos](#objetivos)
- [Funcionalidades](#funcionalidades)
- [Como funciona](#como-funciona)
- [Formato dos modelos](#formato-dos-modelos)
- [Arquitetura](#arquitetura)
- [Tecnologias](#tecnologias)
- [Estrutura do repositório](#estrutura-do-repositório)
- [Como executar](#como-executar)
- [Segurança e privacidade](#segurança-e-privacidade)
- [Requisitos não funcionais](#requisitos-não-funcionais)
- [Histórico de execuções](#histórico-de-execuções)
- [Tratamento de erros](#tratamento-de-erros)
- [Testes e qualidade](#testes-e-qualidade)
- [Escopo do MVP](#escopo-do-mvp)
- [Roadmap](#roadmap)
- [Decisões técnicas](#decisões-técnicas)
- [Como contribuir](#como-contribuir)
- [Licença](#licença)

## Problema e solução

**Problema:** gerar dezenas ou centenas de documentos personalizados (contratos, certificados, cartas) manualmente é lento e sujeito a erros de preenchimento.

**Solução:** o usuário envia uma planilha e um modelo, confere uma prévia e baixa um ZIP com todos os documentos prontos.

**Critério de sucesso do MVP:** subir uma planilha de 100 linhas e um modelo, conferir a prévia e baixar um ZIP com 100 documentos corretos em menos de 1 minuto.

## Objetivos

- Automatizar a geração de documentos personalizados em lote.
- Reduzir tarefas manuais e erros de preenchimento.
- Validar os dados antes de iniciar a geração.
- Permitir a visualização prévia de um documento.
- Disponibilizar todos os documentos em um único arquivo ZIP.
- Registrar o histórico das execuções e os erros encontrados.
- Priorizar privacidade, segurança e processamento sem dependência obrigatória de APIs externas.

## Funcionalidades

### Essenciais (MVP)

- Upload de planilhas Excel (`.xlsx`).
- Upload de modelos Word (`.docx`).
- Identificação automática dos campos (`{{campo}}`) presentes no modelo.
- Sugestão de correspondência entre campos do modelo e colunas da planilha.
- Validação dos dados antes da geração.
- Geração de um documento por linha da planilha.
- Download dos documentos em um arquivo ZIP.
- Tratamento de erros por registro, sem interromper todo o processamento.

### Complementares

- Visualização prévia do primeiro documento.
- Histórico de gerações.
- Relatório de erros por registro.
- Nomes de arquivos configuráveis (ex.: `contrato_{{nome}}.docx`).
- Conversão para PDF.
- Envio de documentos por e-mail.

As funcionalidades complementares serão implementadas conforme as prioridades e a evolução do projeto.

## Como funciona

1. **Envio dos arquivos:** o usuário fornece uma planilha `.xlsx` e um modelo `.docx`.
2. **Leitura:** o sistema identifica as colunas da planilha e os campos do modelo.
3. **Mapeamento:** o sistema sugere a correspondência entre campos e colunas (nomes iguais são pré-mapeados); o usuário pode ajustar.
4. **Validação:** são verificadas colunas ausentes, campos obrigatórios vazios e inconsistências nos dados. Campos sem mapeamento bloqueiam a geração.
5. **Pré-visualização:** o usuário confere um documento de exemplo antes do processamento completo.
6. **Geração em lote:** o modelo é preenchido para cada registro válido.
7. **Empacotamento:** os documentos gerados são reunidos em um ZIP.
8. **Download e histórico:** o usuário baixa o resultado e o sistema registra os metadados da execução.
9. **Limpeza:** os arquivos temporários são removidos conforme a política de retenção.

## Formato dos modelos

Os modelos Word usam marcadores no formato `{{campo}}` para indicar o que será preenchido.

```text
CONTRATO DE PRESTAÇÃO DE SERVIÇOS

Contratante: {{nome}}
CPF: {{cpf}}
Endereço: {{endereco}}
Valor do contrato: {{valor}}
```

A planilha deve ter colunas correspondentes aos campos do modelo:

| nome | cpf | endereco | valor |
|---|---|---|---|
| Ana Silva | 00000000000 | Rua A, 100 | 1500,00 |
| João Souza | 11111111111 | Rua B, 200 | 2200,00 |

O sistema usa os dados de cada linha para gerar um documento individual.

**Atenção ao criar o modelo:** o Word às vezes divide um marcador em vários trechos internos (*runs*) e o campo deixa de ser substituído. Digite cada marcador de uma só vez, sem alterar a formatação no meio dele.

## Arquitetura

O DocBatch adota uma **arquitetura de microsserviços simples**. Cada serviço tem uma responsabilidade específica, expõe uma API REST e pode ser executado e testado de forma independente. A interface web fala apenas com a API de entrada (gateway), que encaminha as solicitações aos serviços de processamento.

```text
                  ┌──────────────────────────────┐
                  │   Interface Web (Streamlit)  │
                  └──────────────┬───────────────┘
                                 │ HTTP / REST
                  ┌──────────────▼───────────────┐
                  │     API / Gateway (FastAPI)  │
                  └──┬─────────┬─────────┬────┬──┘
                     │         │         │    │
        ┌────────────▼──┐ ┌────▼──────┐ ┌▼────▼─────────┐ ┌─────────────────┐
        │ Serviço de    │ │ Serviço de│ │ Serviço de    │ │ Serviço de      │
        │ arquivos      │ │ validação │ │ geração       │ │ empacotamento   │
        └───────────────┘ └───────────┘ └───────────────┘ └─────────────────┘

        ┌──────────────────────┐        ┌──────────────────────────────────┐
        │  SQLite (histórico)  │        │  Armazenamento temporário        │
        └──────────────────────┘        └──────────────────────────────────┘
```

### Componentes

| Componente | Responsabilidade |
|---|---|
| Interface Web (Streamlit) | Upload, mapeamento, validação, prévia e download. |
| API / Gateway (FastAPI) | Entrada única das requisições, validação inicial e encaminhamento aos serviços. |
| Serviço de arquivos | Leitura de planilhas Excel e extração dos campos dos modelos Word. |
| Serviço de validação | Verificação de colunas ausentes, campos vazios e incompatibilidades. |
| Serviço de geração | Preenchimento do modelo e criação de um documento por linha válida. |
| Serviço de empacotamento | Agrupamento dos documentos em ZIP e disponibilização para download. |
| Banco de dados (SQLite) | Histórico e erros das execuções, sem guardar o conteúdo dos documentos. |
| Armazenamento temporário | Arquivos de entrada e resultados apenas durante o processamento. |

### Comunicação entre serviços

- APIs REST com contratos documentados, em `shared/contracts/`.
- Formato de resposta e de erro padronizados em todos os serviços.
- Comunicação direta e simples. Filas, brokers e outras infraestruturas só entram se houver necessidade comprovada.

### Regras arquiteturais

- **Sem APIs externas obrigatórias:** o fluxo principal funciona sem serviços pagos, IA externa ou armazenamento em nuvem.
- **Isolamento de falhas:** um registro com erro não interrompe a geração dos demais.
- **Privacidade:** o conteúdo dos documentos não é persistido no histórico, e os arquivos temporários são removidos ao fim do ciclo de vida.
- **Validação de arquivos:** extensão, tamanho, conteúdo e nomes são verificados antes de processar.
- **Contratos definidos:** serviços trocam dados por APIs documentadas.
- **Testes automatizados:** leitura, mapeamento, validação, geração e empacotamento são cobertos.

## Tecnologias

| Tecnologia | Utilização |
|---|---|
| Python | Linguagem principal de todos os serviços. |
| FastAPI | APIs REST (gateway e serviços). |
| Streamlit | Interface web do MVP. |
| pandas | Leitura e manipulação de dados tabulares. |
| python-docx / docxtpl | Manipulação e preenchimento de documentos Word. |
| SQLite | Persistência do histórico de execuções. |
| zipfile | Criação dos arquivos ZIP. |
| Docker | Isolamento e padronização do ambiente. |
| pytest | Testes automatizados. |
| Ruff | Análise estática e qualidade de código. |
| GitHub Actions | Testes e verificações de integração contínua. |

As dependências definitivas ficam registradas em `requirements.txt`.

## Estrutura do repositório

```text
docbatch/
├── frontend/
│   └── app.py
├── gateway/
│   └── main.py
├── services/
│   ├── file_service/
│   ├── validation_service/
│   ├── generation_service/
│   └── packaging_service/
├── shared/
│   └── contracts/
├── storage/
├── tests/
├── docker-compose.yml
├── requirements.txt
└── README.md
```

| Diretório | Conteúdo |
|---|---|
| `frontend/` | Interface web e interação com o usuário. |
| `gateway/` | API de entrada e encaminhamento das solicitações. |
| `services/` | Serviços de arquivos, validação, geração e empacotamento. |
| `shared/contracts/` | Contratos, esquemas e estruturas compartilhadas entre serviços. |
| `storage/` | Persistência (SQLite) e armazenamento temporário. |
| `tests/` | Testes unitários e de integração, com planilhas e modelos de exemplo. |

A estrutura pode ser ajustada durante a implementação, desde que as responsabilidades continuem bem definidas.

## Como executar

> Comandos previstos. Portas e nomes de módulos podem mudar durante a implementação.

### Pré-requisitos

- Python 3.11 ou superior
- Docker e Docker Compose (para execução conjunta)

### Com Docker (recomendado)

```bash
docker compose up --build
```

Depois, acesse a interface Streamlit no endereço indicado no terminal.

### Localmente, sem Docker

```bash
# 1. Criar e ativar o ambiente virtual
python -m venv .venv
source .venv/bin/activate        # Linux/macOS
# .venv\Scripts\activate         # Windows

# 2. Instalar as dependências
pip install -r requirements.txt

# 3. Subir o gateway e os serviços (um terminal para cada)
uvicorn gateway.main:app --reload --port 8000
# repetir para cada serviço em services/, cada um em uma porta própria

# 4. Subir a interface
streamlit run frontend/app.py
```

### Testes e qualidade

```bash
pytest
ruff check .
```

## Segurança e privacidade

- O fluxo principal **não depende de APIs externas** ou serviços pagos: os documentos nunca saem do ambiente de execução.
- Os arquivos enviados são validados antes do processamento: formato, arquivo corrompido, tamanho do upload e quantidade de registros.
- Nomes de arquivos são sanitizados para evitar nomes maliciosos.
- O conteúdo dos documentos **não** é armazenado permanentemente. O histórico guarda apenas metadados.
- Uploads, documentos gerados e ZIPs são temporários e removidos conforme a política de ciclo de vida (a definir na tarefa de limpeza de temporários).
- Credenciais e segredos nunca entram no código-fonte (use variáveis de ambiente).
- Os logs evitam dados pessoais desnecessários.
- O tratamento de dados pessoais deve observar a **LGPD**.

## Requisitos não funcionais

### Desempenho

- Gerar 100 documentos em menos de 1 minuto.
- Gerar 500 documentos em menos de 2 minutos.
- Exibir informações de andamento durante o processamento.

Esses tempos são metas e serão confirmados por testes em ambiente de referência documentado.

### Confiabilidade

- Erros em registros individuais não interrompem o lote.
- Cada erro é associado ao registro correspondente.
- Arquivos e dados são validados antes da geração.

### Manutenibilidade

- Cada serviço tem responsabilidades delimitadas e pode ser testado isoladamente.
- APIs têm contratos documentados.
- A lógica de negócio é testável independentemente da interface.
- Código com padrões consistentes de organização e estilo.

## Histórico de execuções

Metadados registrados em SQLite:

- Data e hora da execução.
- Modelo e planilha utilizados (nomes).
- Quantidade total de registros processados.
- Quantidade de documentos gerados com sucesso.
- Quantidade de registros com erro.
- Resumo dos erros (número da linha, campo e mensagem).

O histórico **não** guarda o conteúdo das planilhas, dos modelos ou dos documentos gerados.

## Tratamento de erros

O sistema identifica situações como:

- Arquivo em formato não suportado ou corrompido.
- Planilha sem dados ou com estrutura inválida.
- Campo do modelo sem coluna correspondente.
- Dados obrigatórios ausentes.
- Falha ao preencher um documento.
- Falha ao criar o ZIP.
- Falha de comunicação entre serviços.

Os erros apresentam uma descrição clara ao usuário e informações técnicas para diagnóstico, sem expor dados pessoais.

**Exemplo:** ao gerar 100 linhas com 3 inválidas, o ZIP contém 97 documentos e um relatório lista as 3 falhas.

## Testes e qualidade

- Leitura de arquivos Excel.
- Extração dos marcadores do Word (incluindo marcadores quebrados em vários *runs*).
- Mapeamento e validação (campo vazio, coluna faltando, caracteres especiais).
- Geração de documentos.
- Criação e integridade do ZIP.
- Integração entre os serviços.
- Tratamento de falhas e limpeza de arquivos temporários.
- Desempenho para lotes de 100 e 500 documentos.

Meta inicial de cobertura: acima de 70% na lógica central, sem substituir a análise da qualidade dos testes. O GitHub Actions executa as verificações automatizadas.

## Escopo do MVP

### Incluído

- Upload de arquivos `.xlsx` e `.docx`.
- Extração e mapeamento dos campos.
- Validação dos dados.
- Pré-visualização de um documento.
- Geração de documentos em lote.
- Download em ZIP.
- Tratamento de erros por registro.
- Registro básico do histórico.

### Fora do escopo inicial

- Modelos de entrada em PDF.
- Autenticação multiusuário.
- Editor visual de modelos dentro da aplicação.
- Envio de e-mails como requisito obrigatório.
- Dependência de serviços externos para o fluxo principal.

## Roadmap

### Backlog por prioridade

**P0, essenciais**

| Área | Tarefa |
|---|---|
| Arquitetura | Definir arquitetura de microsserviços |
| Arquitetura | Definir contratos de comunicação entre serviços |
| Arquitetura | Definir estrutura de pastas e ambiente Python |
| Segurança | Garantir processamento sem APIs externas obrigatórias |
| Segurança | Implementar validação e limites dos arquivos |
| Segurança | Definir limpeza de arquivos temporários |
| Funcionalidade | Implementar leitura de planilhas Excel |
| Funcionalidade | Implementar extração de placeholders do Word |
| Funcionalidade | Implementar mapeamento e validação de campos |
| Funcionalidade | Implementar geração em lote e exportação ZIP |

**P1, importantes**

| Área | Tarefa |
|---|---|
| Funcionalidade | Implementar preview e relatório de erros |
| Qualidade | Implementar histórico de gerações |
| Qualidade | Criar testes automatizados e integração contínua |
| Qualidade | Medir desempenho e documentar o projeto |

### Etapas de desenvolvimento

**Sprint 1, fundação:** repositório e estrutura, ambiente Python, contratos de comunicação, testes e CI, leitura da planilha e extração dos campos, interface inicial de upload.

**Sprint 2, fluxo principal:** mapeamento, validação, geração de documentos, criação do ZIP, integração da interface com as APIs, teste do fluxo completo.

**Sprint 3, qualidade e finalização:** prévia e relatório de erros, histórico, segurança e limpeza de temporários, testes de desempenho, documentação e implantação.

## Decisões técnicas

### Microsserviços em vez de camadas (DOC-01)

A proposta inicial do projeto previa uma aplicação em camadas (UI, serviços, núcleo e armazenamento). A equipe optou por microsserviços, o que muda a estrutura de pastas, a execução e os testes.

- **Vantagens:** responsabilidades isoladas, serviços executáveis e testáveis de forma independente, base para evolução futura.
- **Custos:** mais complexidade de execução (Docker Compose), de comunicação (contratos HTTP, falhas entre serviços) e de testes de integração.
- **Mitigação:** manter a comunicação simples, sem filas ou brokers, e contratos padronizados em `shared/contracts/`.

Qualquer mudança relevante nessa decisão deve ser registrada e justificada neste documento.

### Outros pontos

| Tema | Decisão ou risco | Mitigação |
|---|---|---|
| Motor de templates | `docxtpl` usa Jinja2 e aceita `{{campo}}` | Usar como base do motor de geração |
| Marcadores quebrados | O Word pode dividir `{{nome}}` em vários *runs* | Documentar a orientação e cobrir nos testes |
| Modelo em PDF | Substituir texto em PDF comum é frágil | MVP aceita só `.docx`; PDF pode ser saída |
| Conversão DOCX para PDF | Requer LibreOffice (`soffice --headless`) | Instalar na imagem Docker |
| Tipos de dados | Datas e valores monetários chegam "crus" do Excel | Formatadores configuráveis (dd/mm/aaaa, R$ 1.234,56) |
| Dados pessoais | Planilhas podem ter CPF, nome etc. | Processar em temporário e apagar após o download |
| Persistência | SQLite pode não ser durável em alguns ambientes de deploy | Reavaliar conforme o ambiente de implantação |

## Como contribuir

1. Crie uma branch para a tarefa (`feat/...`, `fix/...`, `docs/...`).
2. Implemente respeitando as responsabilidades de cada componente.
3. Adicione ou atualize os testes automatizados.
4. Execute as verificações de qualidade (`pytest` e `ruff`) antes de abrir um Pull Request.
5. Documente alterações em contratos, arquitetura ou requisitos.
6. Abra um Pull Request com uma descrição clara das mudanças.

Commits seguem o padrão Conventional Commits (`feat:`, `fix:`, `docs:`).

**Definição de pronto:** funciona, tem teste, passou no CI, foi revisado e está documentado.

## Licença

A licença do projeto será definida antes da distribuição pública do software.

---

**DocBatch: menos trabalho manual, mais consistência na geração de documentos.**

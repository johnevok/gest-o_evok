# ArquiPlanner

Gerenciador pessoal de projetos, criações e agenda (Python + Streamlit).

## Estrutura da pasta
```
arquiplanner_projeto/
├── arquiplanner.py          <- app principal
├── integracoes_google.py    <- Gemini, Google Agenda e Drive
├── hero_nuvem.html          <- busto em nuvem de pontos (tela inicial)
├── requirements.txt
├── .gitignore
└── .streamlit/
    └── secrets.toml         <- chaves (não vai para o GitHub)
```
Os dados ficam em `dados_arquiplanner.json`, criado sozinho na primeira execução.

## 1. Instalar (uma vez)
```
python -m venv .venv
.venv\Scripts\activate          (Windows)   |   source .venv/bin/activate   (Mac/Linux)
pip install -r requirements.txt
```
No VS Code: Ctrl+Shift+P > "Python: Select Interpreter" > escolha o `.venv`.

## 2. Rodar
```
python -m streamlit run arquiplanner.py
```

## 3. Ativar o Gemini
Crie a chave em https://aistudio.google.com/apikey e cole em `GEMINI_API_KEY` no `secrets.toml`
(ou na barra lateral do app, em Configurações).

## 4. Ativar Google Agenda e Drive (opcional)
1. https://console.cloud.google.com > crie um projeto.
2. Biblioteca de APIs: ative **Google Calendar API** e **Google Drive API**.
3. IAM > Contas de serviço > criar; depois Chaves > Adicionar chave > JSON (baixa um arquivo).
4. Copie os campos do JSON para a seção `[gcp_service_account]` do `secrets.toml` e descomente-a.
5. **Agenda:** Google Agenda > Configurações da agenda > Compartilhar com pessoas específicas > adicione o
   `client_email` com "Fazer alterações em eventos". Copie o ID da agenda para `GOOGLE_CALENDAR_ID`.
6. **Drive (pastas):** compartilhe sua pasta de projetos com o mesmo e-mail como Editor.
7. **Drive (backup):** crie no Drive um arquivo de texto vazio `dados_arquiplanner.json`, compartilhe
   com o mesmo e-mail como Editor e copie o ID dele para `DRIVE_BACKUP_FILE_ID`.

## Streamlit Cloud
Suba os arquivos (exceto `secrets.toml` e `dados_arquiplanner.json`) para o GitHub e cole o conteúdo do
`secrets.toml` em Settings > Secrets. Com o backup no Drive configurado, os dados são restaurados após cada reinício.

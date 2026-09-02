# Escala do Pátio — JBS Terminais

Quadro de escala operacional: importa as planilhas do mês, monta o turno e
espelha numa televisão.

---

## O que você precisa

Um computador com **Node.js 18 ou mais novo**. Baixe em <https://nodejs.org>
(pegue a versão "LTS"). Serve Windows, Mac ou Linux.

Nada mais. O servidor não usa banco de dados nem biblioteca externa.

---

## Instalar (uma vez só)

Abra o Prompt de Comando (Windows) ou o Terminal, entre na pasta do projeto e
rode estas duas linhas:

```
npm install
npm run build
```

A primeira baixa o que o app precisa para ser montado. A segunda gera a pasta
`dist/`, que é o app pronto. Demora um ou dois minutos.

---

## Usar no dia a dia

```
npm start
```

Aparece algo assim:

```
Escala do Pátio rodando.
  Nesta máquina:  http://localhost:8080
  Na rede (eth0):  http://192.168.1.40:8080

Dados em: .../dados/quadro.json
Para parar, aperte Ctrl+C.
```

- **No computador da sala:** abra `http://localhost:8080`
- **Na televisão:** abra o endereço da linha "Na rede" e clique em **Modo TV**
- **Em qualquer outro computador do terminal:** o mesmo endereço da rede

Todos compartilham o mesmo quadro. O que um altera aparece nos outros em
poucos segundos.

> O computador que roda `npm start` precisa ficar ligado. Se ele desligar, a TV
> para de atualizar até religar.

---

## Onde ficam os dados

Tudo em `dados/quadro.json`, dentro da pasta do projeto.

- É gravado a cada alteração, de forma segura: se a máquina desligar no meio,
  o arquivo não corrompe.
- Uma cópia por dia vai para `dados/backups/`, e as 30 últimas são mantidas.
- **Para fazer backup, copie a pasta `dados/` inteira.** É só isso.
- Para restaurar, pare o servidor, substitua `dados/quadro.json` pela cópia e
  inicie de novo.

---

## Senha

Dentro do app, o botão à direita da barra de ações ativa a proteção. Ao ligar,
todo o conteúdo passa a ser gravado cifrado — quem abrir o arquivo
`quadro.json` no bloco de notas vê texto embaralhado, não nomes e matrículas.

Cada pessoa tem a própria senha e o nome dela fica no histórico.

**Não existe recuperação de senha.** Os dados são cifrados com ela. Se todas se
perderem, o quadro precisa ser recomeçado. Anote em lugar seguro.

---

## Ligar sozinho quando a máquina inicia

**Windows.** Crie um arquivo `escala.bat` na pasta do projeto com:

```bat
@echo off
cd /d "%~dp0"
node servidor.js
```

Aperte `Win+R`, digite `shell:startup` e coloque um atalho desse `.bat` ali.

**Linux.** Crie `/etc/systemd/system/escala.service`:

```ini
[Unit]
Description=Escala do Patio
After=network.target

[Service]
WorkingDirectory=/caminho/para/escala-patio
ExecStart=/usr/bin/node servidor.js
Restart=always
User=SEU_USUARIO

[Install]
WantedBy=multi-user.target
```

Depois: `sudo systemctl enable --now escala`

---

## Trocar a porta

Se a 8080 já estiver ocupada:

```
PORTA=9000 npm start          (Linux e Mac)
set PORTA=9000 && npm start   (Windows)
```

---

## Se der problema

**"npm não é reconhecido"** — o Node.js não está instalado, ou o Prompt foi
aberto antes da instalação. Feche e abra de novo.

**"A pasta dist/ não existe"** — falta rodar `npm run build`.

**A TV não abre o endereço da rede** — o firewall do computador está bloqueando
a porta 8080. Libere-a, ou peça ao TI.

**"EADDRINUSE"** — já tem um servidor rodando. Feche a outra janela.

**A TV não atualiza** — confira se o computador que roda o servidor está ligado
e se a TV está no mesmo endereço.

---

## Para o TI

- Node.js puro, sem dependências no servidor. As dependências (React, Vite,
  SheetJS) são apenas de build e não vão para produção.
- Escuta HTTP na porta 8080. Sem TLS: pensado para rede interna. Para expor
  fora, coloque atrás de um proxy reverso com HTTPS.
- Sem autenticação no nível do servidor. O controle de acesso é a criptografia
  do próprio app, feita no navegador (AES-GCM, chave derivada por PBKDF2 com
  250 mil iterações, envelope por usuário).
- Estado inteiro em um arquivo JSON. Adequado para dezenas de milhares de
  registros; se um dia crescer muito, trocar por SQLite é direto.
- Limite de 6 MB por registro gravado.

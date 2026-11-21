# Leitor QR (web app para iPhone)

App que lê QR codes **a partir de imagens** (fotos da galeria, prints, arquivos).
Roda no Safari do iPhone e pode ser instalado na Tela de Início como um app.
Todo o processamento acontece no aparelho, sem enviar a imagem para nenhum servidor,
e funciona offline depois da primeira abertura.

## Funcionalidades
- Escolher imagem da Fototeca, tirar foto ou selecionar em Arquivos
- Colar imagem da área de transferência
- Detecta QR em fotos grandes, girados e com cores invertidas
- Botões Copiar, Abrir link (para URLs, e-mail, telefone) e Compartilhar
- Histórico das últimas 30 leituras (salvo no aparelho)

## Como colocar no iPhone
O iPhone precisa acessar o app por **HTTPS**. O jeito mais simples é o GitHub Pages:

1. No GitHub: *Settings → Pages → Build and deployment → Deploy from a branch*,
   escolha o branch e a pasta raiz (`/`), e salve.
2. Abra no Safari do iPhone: `https://<usuario>.github.io/<repositorio>/qrcode-app/`
3. Toque em **Compartilhar → Adicionar à Tela de Início**.

Qualquer hospedagem estática com HTTPS serve (Netlify, Vercel, Cloudflare Pages):
basta publicar a pasta `qrcode-app/`.

## Testar no computador
```bash
cd qrcode-app
python3 -m http.server 8000
# abra http://localhost:8000
```

## Arquivos
| Arquivo | Função |
|---|---|
| `index.html` | Interface |
| `app.js` | Leitura da imagem, decodificação e histórico |
| `jsQR.js` | Biblioteca de decodificação [jsQR](https://github.com/cozmo/jsQR) 1.4.0 (Apache-2.0), incluída localmente para funcionar offline |
| `sw.js` | Service worker (cache offline) |
| `manifest.webmanifest`, `icons/` | Instalação na Tela de Início |

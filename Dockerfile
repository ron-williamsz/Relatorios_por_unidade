# Relatórios de Consumo de Água (Zangari) — site estático servido por nginx.
# O conteúdo de relatorio/ (HTML + JS de dados + bibliotecas + logos) já é
# self-contained; aqui só empacotamos num nginx leve com gzip.

FROM nginx:1.27-alpine

# Configuração (gzip para os JS grandes: echarts, jspdf, dados_*.js)
COPY nginx.conf /etc/nginx/conf.d/default.conf

# Arquivos do relatório
COPY relatorio/ /usr/share/nginx/html/

EXPOSE 80

# healthcheck simples
HEALTHCHECK --interval=30s --timeout=3s --retries=3 \
  CMD wget -qO- http://localhost/ >/dev/null 2>&1 || exit 1

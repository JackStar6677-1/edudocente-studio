#!/bin/bash
set -e

echo "🚀 Iniciando despliegue de EduDocente-Studio en Star Server (192.168.0.120)..."

SRC_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"

# 1. Transferir archivos al directorio temporal en Star
echo "📦 Transfiriendo archivos de EduDocente a Star..."
ssh star "mkdir -p /tmp/edudocente_deploy"
rsync -avz --exclude='.git' --exclude='__pycache__' --exclude='output/*' "$SRC_DIR/" star:/tmp/edudocente_deploy/

# 2. Despliegue en el perfil institucional en Star
echo "⚙️ Configurando directorio institucional /home/admin-colegio/edudocente..."
ssh star bash -s << 'REMOTE_SCRIPT'
set -e

sudo mkdir -p /home/admin-colegio/edudocente
sudo mkdir -p /home/admin-colegio/edudocente/output
sudo cp -r /tmp/edudocente_deploy/* /home/admin-colegio/edudocente/
sudo rm -rf /tmp/edudocente_deploy

# Asignar propiedad a admin-colegio y grupo institucional
sudo chown -R admin-colegio:star-colegio-admins /home/admin-colegio/edudocente
sudo chmod -R 775 /home/admin-colegio/edudocente

# Enlace simbolico institucional /srv/colegio/edudocente
sudo mkdir -p /srv/colegio
sudo ln -sfn /home/admin-colegio/edudocente /srv/colegio/edudocente

# 3. Servicio Systemd
echo "🔧 Instalando servicio edudocente.service..."
sudo cp /home/admin-colegio/edudocente/systemd/edudocente.service /etc/systemd/system/
sudo systemctl daemon-reload
sudo systemctl enable edudocente.service
sudo systemctl restart edudocente.service

# 4. Configurar Nginx para estudio.castelgandolfo
echo "🌐 Configurando Nginx Reverse Proxy para EduDocente..."
sudo tee /etc/nginx/sites-available/edudocente > /dev/null << 'NGINX_CONF'
server {
    listen 80;
    server_name estudio.castelgandolfo edudocente.castelgandolfo;

    location / {
        proxy_pass http://127.0.0.1:8086;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
        client_max_body_size 50M;
        proxy_read_timeout 300;
        proxy_connect_timeout 300;
    }
}
NGINX_CONF

sudo ln -sfn /etc/nginx/sites-available/edudocente /etc/nginx/sites-enabled/edudocente
sudo nginx -t
sudo systemctl reload nginx

# 5. Configurar DNS Local en dnsmasq para resolver estudio.castelgandolfo
echo "🏷️ Actualizando resolución DNS local en dnsmasq..."
if [ -f /etc/dnsmasq.d/hosts-locales.conf ]; then
    if ! grep -q "estudio.castelgandolfo" /etc/dnsmasq.d/hosts-locales.conf; then
        echo "address=/estudio.castelgandolfo/192.168.0.120" | sudo tee -a /etc/dnsmasq.d/hosts-locales.conf > /dev/null
        echo "address=/edudocente.castelgandolfo/192.168.0.120" | sudo tee -a /etc/dnsmasq.d/hosts-locales.conf > /dev/null
        sudo systemctl restart dnsmasq 2>/dev/null || sudo systemctl reload dnsmasq 2>/dev/null || true
    fi
fi

# 6. Verificación de Salud
echo "🩺 Verificando estado del servicio..."
sleep 2
systemctl is-active --quiet edudocente.service && echo "✅ Servicio edudocente.service ACTIVO (Running)" || (echo "❌ Error en servicio" && journalctl -u edudocente.service -n 20 --no-pager && exit 1)

curl -sI http://127.0.0.1:8086/ | head -n 1
echo "🎉 Despliegue de EduDocente-Studio completado exitosamente!"
REMOTE_SCRIPT

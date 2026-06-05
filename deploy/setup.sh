#!/bin/bash
# ============================================================
# 美的业务关务看板 - 阿里云 ECS 一键部署脚本
# 适用系统: Ubuntu 20.04+ / Debian 11+ / Alibaba Cloud Linux
#
# 使用方法（在 ECS 上以 root 执行）:
#   curl -O https://raw.githubusercontent.com/cleverguo1979/midea-customs-dashboard/main/deploy/setup.sh
#   bash setup.sh
# ============================================================
set -e

# ---- 配置 ----
DOMAIN="cleverguo.cn"
APP_DIR="/opt/midea-dashboard"
APP_USER="midea"
GUNICORN_PORT="8888"
GUNICORN_WORKERS="2"

echo "============================================"
echo "  美的关务看板 - 阿里云部署脚本"
echo "  域名: ${DOMAIN}"
echo "============================================"

# ---- 1. 检测系统 ----
if [ -f /etc/os-release ]; then
    . /etc/os-release
    OS=$ID
else
    echo "无法检测系统版本，请手动安装"
    exit 1
fi

echo "[1/7] 检测到系统: ${OS}"

# ---- 2. 安装依赖 ----
echo "[2/7] 安装系统依赖..."

if [ "$OS" = "ubuntu" ] || [ "$OS" = "debian" ]; then
    apt update -qq
    apt install -y -qq python3 python3-pip nginx git certbot python3-certbot-nginx curl
elif [ "$OS" = "alinux" ] || [ "$OS" = "centos" ] || [ "$OS" = "rhel" ]; then
    yum install -y python3 python3-pip nginx git curl
    # CentOS 7 可能需要 EPEL 来装 certbot
    yum install -y epel-release 2>/dev/null || true
    yum install -y certbot python3-certbot-nginx 2>/dev/null || true
else
    echo "不支持的系统: ${OS}，请手动安装 python3, nginx, git"
    exit 1
fi

# ---- 3. 创建应用用户和目录 ----
echo "[3/7] 创建应用目录..."

if ! id -u ${APP_USER} >/dev/null 2>&1; then
    useradd -r -s /bin/false ${APP_USER}
fi

mkdir -p ${APP_DIR}
cd ${APP_DIR}

# ---- 4. 拉取代码 ----
echo "[4/7] 拉取最新代码..."

if [ -d "${APP_DIR}/.git" ]; then
    cd ${APP_DIR}
    git pull origin main
else
    git clone https://github.com/cleverguo1979/midea-customs-dashboard.git ${APP_DIR}
    cd ${APP_DIR}
fi

# ---- 5. 安装 Python 依赖 ----
echo "[5/7] 安装 Python 依赖..."

pip3 install flask openpyxl gunicorn

# ---- 6. 配置 systemd 服务 ----
echo "[6/7] 配置系统服务..."

cat > /etc/systemd/system/midea-dashboard.service << SYSTEMDEOF
[Unit]
Description=Midea Customs Dashboard
After=network.target

[Service]
Type=simple
User=${APP_USER}
Group=${APP_USER}
WorkingDirectory=${APP_DIR}
ExecStart=/usr/local/bin/gunicorn -w ${GUNICORN_WORKERS} -b 127.0.0.1:${GUNICORN_PORT} server:app
Restart=always
RestartSec=5
Environment=PYTHONUNBUFFERED=1

[Install]
WantedBy=multi-user.target
SYSTEMDEOF

# 确保数据库文件权限正确
chown -R ${APP_USER}:${APP_USER} ${APP_DIR}

systemctl daemon-reload
systemctl enable midea-dashboard
systemctl restart midea-dashboard

# ---- 7. 配置 Nginx ----
echo "[7/7] 配置 Nginx 反向代理..."

cat > /etc/nginx/sites-available/midea-dashboard << NGINXEOF
# HTTP → 先跑起来，后面再配 SSL
server {
    listen 80;
    server_name ${DOMAIN};

    # 日志
    access_log /var/log/nginx/midea-access.log;
    error_log  /var/log/nginx/midea-error.log;

    # 上传大小限制
    client_max_body_size 50m;

    # 静态文件 — 直接由 nginx 提供
    location / {
        root ${APP_DIR};
        try_files \$uri @gunicorn;
    }

    # API 请求 — 转发给 gunicorn
    location /api/ {
        proxy_pass http://127.0.0.1:${GUNICORN_PORT};
        proxy_set_header Host \$host;
        proxy_set_header X-Real-IP \$remote_addr;
        proxy_set_header X-Forwarded-For \$proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto \$scheme;
        proxy_read_timeout 120s;
    }

    # 非文件请求（SPA fallback）— 返回 index.html
    location @gunicorn {
        proxy_pass http://127.0.0.1:${GUNICORN_PORT};
        proxy_set_header Host \$host;
        proxy_set_header X-Real-IP \$remote_addr;
        proxy_set_header X-Forwarded-For \$proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto \$scheme;
    }
}
NGINXEOF

# 对于 Ubuntu/Debian
if [ -d /etc/nginx/sites-enabled ]; then
    ln -sf /etc/nginx/sites-available/midea-dashboard /etc/nginx/sites-enabled/
    # 删除默认站点
    rm -f /etc/nginx/sites-enabled/default
# 对于 CentOS/Alibaba Cloud Linux
elif [ -d /etc/nginx/conf.d ]; then
    ln -sf /etc/nginx/sites-available/midea-dashboard /etc/nginx/conf.d/midea-dashboard.conf
fi

# 测试 nginx 配置
nginx -t

# 重载 nginx
systemctl enable nginx
systemctl restart nginx

# ---- 完成 ----
echo ""
echo "============================================"
echo "  ✅ 部署完成！"
echo "============================================"
echo ""
echo "  访问地址: http://${DOMAIN}"
echo "  应用目录: ${APP_DIR}"
echo ""
echo "  常用命令:"
echo "    systemctl status midea-dashboard   # 查看后端状态"
echo "    systemctl restart midea-dashboard  # 重启后端"
echo "    journalctl -u midea-dashboard -f   # 查看后端日志"
echo "    nginx -t && nginx -s reload        # 重载 nginx"
echo ""
echo "  下一步 — 配置 SSL 证书（推荐）:"
echo "    certbot --nginx -d ${DOMAIN}"
echo ""
echo "============================================"

# 检查服务状态
echo ""
echo "服务状态检查:"
systemctl status midea-dashboard --no-pager 2>/dev/null | head -10 || true

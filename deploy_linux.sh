#!/bin/bash
# =============================================================
# 智能油库日报机器人 —— Linux 一键部署脚本
# 适用：Ubuntu 20.04+ / CentOS 7+ (systemd)
# 用法：chmod +x deploy_linux.sh && sudo ./deploy_linux.sh
# =============================================================
set -e

INSTALL_DIR="/opt/daily_report_robot"
SERVICE="daily-report-robot"
PYTHON=$(which python3)

echo ">>> [1/5] 创建安装目录 $INSTALL_DIR"
mkdir -p "$INSTALL_DIR"

echo ">>> [2/5] 复制程序文件"
cp main.py          "$INSTALL_DIR/"
cp requirements.txt "$INSTALL_DIR/"
# 本地配置文件（若已配置在线表格可忽略）
[ -f device_config.xlsx ] && cp device_config.xlsx "$INSTALL_DIR/" || true

echo ">>> [3/5] 安装 Python 依赖"
pip3 install -r "$INSTALL_DIR/requirements.txt"

echo ">>> [4/5] 创建 systemd 服务"
cat > /etc/systemd/system/$SERVICE.service << EOF
[Unit]
Description=Daily Report Robot (智能油库日报机器人)
After=network-online.target
Wants=network-online.target

[Service]
Type=simple
User=root
WorkingDirectory=$INSTALL_DIR

# ── 敏感凭据建议通过环境变量注入，避免写在代码中 ──
# Environment="DB_PASSWORD=your_password"
# Environment="WEBHOOK_KEY=your_key"
# Environment="FEISHU_APP_ID=cli_xxx"
# Environment="FEISHU_APP_SECRET=xxx"

ExecStart=$PYTHON $INSTALL_DIR/main.py
Restart=always
RestartSec=30
StandardOutput=journal
StandardError=journal

[Install]
WantedBy=multi-user.target
EOF

echo ">>> [5/5] 启动服务并设置开机自启"
systemctl daemon-reload
systemctl enable  $SERVICE
systemctl restart $SERVICE

echo ""
echo "=============================="
echo "  部署完成！"
echo "=============================="
echo "  查看运行状态：systemctl status $SERVICE"
echo "  实时查看日志：journalctl -u $SERVICE -f"
echo "  停止服务：    systemctl stop $SERVICE"
echo "  重启服务：    systemctl restart $SERVICE"
echo "=============================="
